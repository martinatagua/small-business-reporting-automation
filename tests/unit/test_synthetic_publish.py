"""T3-17: protocolo de escritura y publicación (Plan §8.3). Todo en ``tmp_path``; fallos inyectados mediante los
puntos ``synthetic._fs_*``."""

import errno
import hashlib
import json
import os
import shutil
from pathlib import Path

import pytest

from small_business_reporting_automation import synthetic as syn

import sample_oracle as oracle

ALL = oracle.EXPECTED_MANIFEST_PATHS + ["ground_truth/manifest.json"]


def snapshot(root: Path) -> dict:
    return {rel: hashlib.sha256((root / rel).read_bytes()).hexdigest() for rel in ALL if (root / rel).exists()}


def consistent(root: Path) -> dict:
    raw = (root / "ground_truth/manifest.json").read_bytes()
    manifest = json.loads(raw.decode("utf-8"))
    assert oracle.manifest_problems(manifest) == []
    for entry in manifest["files"]:
        assert oracle.file_digest(root, entry["path"]) == entry["content_sha256"]
    return manifest


def fails(kind, **kw):
    with pytest.raises(syn.GenerateError) as exc:
        syn.generate(**kw)
    assert exc.value.kind == kind, exc.value.message
    return exc.value.message


@pytest.fixture()
def old(tmp_path):
    root = tmp_path / "data"
    syn.generate(1, root)
    return root


def staging(root):
    return root / syn.STAGING_NAME


# ---- directorios y tipos de ruta (C5, N-1) --------------------------------------------------------------


def test_output_root_absent_with_existing_parent_is_created(tmp_path):
    root = tmp_path / "nuevo"
    syn.generate(5, root)
    assert consistent(root)["seed"] == 5 and not staging(root).exists()


def test_output_root_absent_without_parent_creates_nothing(tmp_path):
    fails("prepare", seed=5, output_root=tmp_path / "a" / "b")
    assert not (tmp_path / "a").exists()


def test_output_root_incompatible(tmp_path):
    (tmp_path / "f").write_text("x")
    fails("prepare", seed=5, output_root=tmp_path / "f")
    assert (tmp_path / "f").read_text() == "x"
    (tmp_path / "real").mkdir()
    os.symlink(tmp_path / "real", tmp_path / "link")
    fails("prepare", seed=5, output_root=tmp_path / "link")
    assert list((tmp_path / "real").iterdir()) == []


def test_missing_subdirectories_are_created_after_staging(tmp_path, monkeypatch):
    root = tmp_path / "data"
    root.mkdir()
    calls = []
    real = syn._fs_mkdir
    monkeypatch.setattr(syn, "_fs_mkdir", lambda p: (calls.append(Path(p)), real(p))[1])
    syn.generate(5, root)
    consistent(root)
    stg = calls.index(staging(root))
    first_dest = min(calls.index(root / s) for s in syn.SUBDIRS)
    assert stg < first_dest and all(p.parent == staging(root) for p in calls[stg + 1:stg + 4])


def test_incompatible_subdirectory_or_target(tmp_path):
    root = tmp_path / "data"
    root.mkdir()
    (root / "raw").write_text("no soy un directorio")
    fails("prepare", seed=5, output_root=root)
    assert (root / "raw").read_text() == "no soy un directorio" and not staging(root).exists()
    (root / "raw").unlink()
    (tmp_path / "otro").mkdir()
    os.symlink(tmp_path / "otro", root / "raw")
    fails("prepare", seed=5, output_root=root)
    (root / "raw").unlink()
    # rutas de destino: directorio o enlace donde va un archivo
    syn.generate(5, root)
    (root / "reference/products.csv").unlink()
    (root / "reference/products.csv").mkdir()
    fails("prepare", seed=6, output_root=root, overwrite=True)
    (root / "reference/products.csv").rmdir()
    os.symlink(root / "raw/centro_ventas_2026-10.csv", root / "reference/products.csv")
    before = snapshot(root)
    fails("prepare", seed=6, output_root=root, overwrite=True)
    assert snapshot(root) == before and not staging(root).exists()


def test_existing_target_without_overwrite(old):
    before = snapshot(old)
    fails("prepare", seed=2, output_root=old)
    assert snapshot(old) == before and not staging(old).exists()


def test_directory_preparation_failure_keeps_previous_set_and_manifest(old, monkeypatch):
    before = snapshot(old)
    monkeypatch.setattr(syn, "_fs_dev", lambda p: 1 if Path(p) == staging(old) else 2)  # otro sistema de archivos
    fails("prepare", seed=2, output_root=old, overwrite=True)
    assert snapshot(old) == before and "ground_truth/manifest.json" in before and not staging(old).exists()
    consistent(old)


def test_directory_failure_removes_directories_created_by_the_run(tmp_path, monkeypatch):
    root = tmp_path / "nuevo"
    monkeypatch.setattr(syn, "_fs_dev", lambda p: 1 if Path(p) == staging(root) else 2)
    fails("prepare", seed=2, output_root=root)
    assert not root.exists()  # la raíz la creó la ejecución y quedó vacía
    root.mkdir()
    fails("prepare", seed=2, output_root=root)
    assert list(root.iterdir()) == []  # solo se eliminan los subdirectorios creados por la ejecución


def test_mkdir_failure_in_3b_keeps_previous_set(old, monkeypatch):
    shutil.rmtree(old / "reference")  # fuerza la creación de un subdirectorio en 3b
    before = snapshot(old)
    real = syn._fs_mkdir

    def boom(path):
        if Path(path) == old / "reference":
            raise OSError(errno.EACCES, "inyectado")
        real(path)

    monkeypatch.setattr(syn, "_fs_mkdir", boom)
    fails("prepare", seed=2, output_root=old, overwrite=True)
    assert snapshot(old) == before and not staging(old).exists()
    assert (old / "ground_truth/manifest.json").exists()


# ---- staging residual (C2) ---------------------------------------------------------------------------------


def test_residual_staging_blocks_without_deleting(old):
    before = snapshot(old)
    staging(old).mkdir()
    (staging(old) / "marca.txt").write_text("ajeno")
    fails("prepare", seed=1, output_root=old, overwrite=True)
    assert (staging(old) / "marca.txt").read_text() == "ajeno" and snapshot(old) == before
    shutil.rmtree(staging(old))  # el propietario lo elimina a mano
    syn.generate(1, old, overwrite=True)
    consistent(old)


# ---- fase 3: fallo de escritura/validación ----------------------------------------------------------------


def test_phase3_failure_leaves_destination_intact(old, monkeypatch):
    before = snapshot(old)
    real = syn._fs_write
    count = {"n": 0}

    def boom(path, data):
        count["n"] += 1
        if count["n"] == 4:
            raise OSError(errno.ENOSPC, "inyectado")
        real(path, data)

    monkeypatch.setattr(syn, "_fs_write", boom)
    fails("prepare", seed=2, output_root=old, overwrite=True)
    assert snapshot(old) == before and not staging(old).exists()


def test_phase3_validation_failure(old, monkeypatch):
    before = snapshot(old)
    real = syn._fs_write

    def corrupt(path, data):
        if str(path).endswith("products.csv"):
            data = data.replace(b"\n", b"\r\n")  # cambia los terminadores sin alterar los registros
        real(path, data)

    monkeypatch.setattr(syn, "_fs_write", corrupt)
    msg = fails("prepare", seed=2, output_root=old, overwrite=True)
    assert "terminador" in msg and snapshot(old) == before and not staging(old).exists()


# ---- fase 4 ---------------------------------------------------------------------------------------------------


@pytest.mark.parametrize("k", [1, 5, 9, 10])
def test_phase4_replace_failure_leaves_no_manifest(old, monkeypatch, k):
    real = syn._fs_replace
    calls = []

    def boom(src, dst):
        calls.append(Path(dst))
        if len(calls) == k:
            raise OSError(errno.EIO, "inyectado")
        real(src, dst)

    monkeypatch.setattr(syn, "_fs_replace", boom)
    msg = fails("interrupted", seed=2, output_root=old, overwrite=True)
    assert "manifest.json ausente" in msg and "--overwrite" in msg
    assert not (old / "ground_truth/manifest.json").exists() and len(calls) == k  # no se sigue publicando
    assert not staging(old).exists()
    syn.generate(2, old, overwrite=True)  # recuperación: misma semilla con --overwrite
    assert consistent(old)["seed"] == 2


def test_phase4_exdev_stops_without_copying(old, monkeypatch):
    def boom(src, dst):
        raise OSError(errno.EXDEV, "Invalid cross-device link")

    monkeypatch.setattr(syn, "_fs_replace", boom)
    msg = fails("interrupted", seed=2, output_root=old, overwrite=True)
    assert "manifest.json ausente" in msg and not (old / "ground_truth/manifest.json").exists()
    # sin copia alternativa: los archivos de destino conservan el contenido anterior
    assert oracle.file_digest(old, "reference/products.csv")


def test_phase4a_failure_keeps_previous_set(old, monkeypatch):
    before = snapshot(old)

    def boom(path):
        raise OSError(errno.EPERM, "inyectado")

    monkeypatch.setattr(syn, "_fs_unlink", boom)
    msg = fails("prepare", seed=2, output_root=old, overwrite=True)
    assert "conjunto anterior sigue intacto" in msg and snapshot(old) == before


def test_cleanup_failure_after_manifest_publication(old, monkeypatch):
    real = syn._fs_rmdir

    def boom(path):
        if Path(path) == staging(old):
            raise OSError(errno.EBUSY, "inyectado")
        real(path)

    monkeypatch.setattr(syn, "_fs_rmdir", boom)
    msg = fails("cleanup", seed=2, output_root=old, overwrite=True)
    assert "conjunto publicado y validado" in msg and "manifest.json presente" in msg
    assert "ausente" not in msg and "incompleto:" not in msg
    assert consistent(old)["seed"] == 2  # el conjunto nuevo es consistente y el manifiesto sigue publicado
    assert staging(old).is_dir()  # residuo: no se borra automáticamente
    monkeypatch.undo()
    fails("prepare", seed=2, output_root=old, overwrite=True)  # el residuo bloquea la siguiente ejecución
    assert staging(old).is_dir()
    shutil.rmtree(staging(old))
    syn.generate(2, old, overwrite=True)
    assert not staging(old).exists()


def test_publication_order_and_overwrite_cycle(old, monkeypatch):
    log = []
    ru, rr = syn._fs_unlink, syn._fs_replace
    monkeypatch.setattr(syn, "_fs_unlink", lambda p: (log.append(("unlink", Path(p))), ru(p))[1])
    monkeypatch.setattr(syn, "_fs_replace", lambda s, d: (log.append(("replace", Path(d))), rr(s, d))[1])
    syn.generate(3, old, overwrite=True)
    assert log[0] == ("unlink", old / "ground_truth/manifest.json")  # el manifiesto anterior se elimina primero
    assert [d.relative_to(old).as_posix() for _op, d in log[1:]] == oracle.EXPECTED_MANIFEST_PATHS + [
        "ground_truth/manifest.json"]  # los nueve en orden canónico y el manifiesto el último
    assert consistent(old)["seed"] == 3 and not staging(old).exists()
    # los .gitkeep y archivos ajenos no se tocan
    (old / "raw/.gitkeep").write_text("")
    (old / "raw/ajeno.txt").write_text("x")
    syn.generate(3, old, overwrite=True)
    assert (old / "raw/.gitkeep").exists() and (old / "raw/ajeno.txt").read_text() == "x"
