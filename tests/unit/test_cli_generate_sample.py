"""T3-15: ``sb-report generate-sample`` (Plan §8.1, §8.2)."""

import errno
import json
import re
import shutil
from pathlib import Path

import pytest

from small_business_reporting_automation import cli
from small_business_reporting_automation import synthetic as syn

REPO = Path(__file__).resolve().parents[2]
DEMO_TEXT = (REPO / "config" / "demo.toml").read_text(encoding="utf-8")


def run(capsys, *argv):
    code = cli.main(list(argv))
    captured = capsys.readouterr()
    return code, captured.out, captured.err


def test_seed_without_any_configuration(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)  # sin config/demo.toml en el directorio actual
    code, out, err = run(capsys, "generate-sample", "--seed", "5")
    assert code == 0 and "177 registros" in out and err == ""
    manifest = json.loads((tmp_path / "data/ground_truth/manifest.json").read_text(encoding="utf-8"))
    assert manifest["seed"] == 5  # --output-root por defecto: data (creada: N-1)


def test_seed_from_demo_config(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "config").mkdir()
    (tmp_path / "config/demo.toml").write_text(DEMO_TEXT, encoding="utf-8")
    code, out, _ = run(capsys, "generate-sample", "--output-root", "salida")
    assert code == 0 and "semilla 20261001" in out
    assert json.loads((tmp_path / "salida/ground_truth/manifest.json").read_text())["seed"] == 20261001


def test_seed_prevails_over_configuration(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "config").mkdir()
    (tmp_path / "config/demo.toml").write_text("esto no es TOML válido [", encoding="utf-8")  # no se lee
    code, out, _ = run(capsys, "generate-sample", "--seed", "9", "--output-root", "o")
    assert code == 0 and "semilla 9" in out


def test_config_problems_exit_2(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    code, _, err = run(capsys, "generate-sample")  # archivo ausente
    assert code == 2 and "C01_INVALID_CONFIG" in err
    (tmp_path / "config").mkdir()
    (tmp_path / "config/demo.toml").write_text("[[[", encoding="utf-8")
    code, _, err = run(capsys, "generate-sample")  # C01
    assert code == 2 and "C01_INVALID_CONFIG" in err
    no_sample = re.sub(r"\[sample\]\nseed = \d+[^\n]*\n", "", DEMO_TEXT)
    assert "[sample]" not in no_sample.replace("[sample]:", "")
    (tmp_path / "config/demo.toml").write_text(no_sample, encoding="utf-8")
    code, _, err = run(capsys, "generate-sample")  # válida pero sin [sample]
    assert code == 2 and "no hay semilla" in err and not (tmp_path / "data").exists()


@pytest.mark.parametrize("bad", ["-1", "9223372036854775808", "abc", "1.5", ""])
def test_invalid_seed_exits_2(bad, capsys):
    with pytest.raises(SystemExit) as exc:
        cli.main(["generate-sample", "--seed", bad])
    assert exc.value.code == 2
    capsys.readouterr()


def test_max_seed_is_accepted(tmp_path, capsys):
    code, _, _ = run(capsys, "generate-sample", "--seed", str(2**63 - 1), "--output-root", str(tmp_path / "d"))
    assert code == 0


@pytest.mark.parametrize("argv", [["generate-sample", "--config", "x.toml"], ["run"], ["recover"],
                                  ["generate-sample", "--bogus"]])
def test_no_config_run_or_recover(argv, capsys):
    with pytest.raises(SystemExit) as exc:
        cli.main(argv)
    assert exc.value.code == 2
    capsys.readouterr()


def test_help(capsys):
    with pytest.raises(SystemExit) as exc:
        cli.main(["--help"])
    assert exc.value.code == 0
    out = capsys.readouterr().out
    assert "generate-sample" in out and "recover" not in out and "run" not in out.replace("sb-report", "")
    with pytest.raises(SystemExit):
        cli.main(["generate-sample", "--help"])
    out = capsys.readouterr().out
    assert all(opt in out for opt in ("--seed", "--output-root", "--overwrite")) and "--config" not in out


def test_existing_target_without_overwrite_and_with(tmp_path, capsys):
    root = str(tmp_path / "d")
    assert run(capsys, "generate-sample", "--seed", "1", "--output-root", root)[0] == 0
    code, _, err = run(capsys, "generate-sample", "--seed", "1", "--output-root", root)
    assert code == 2 and "--overwrite" in err
    assert run(capsys, "generate-sample", "--seed", "2", "--output-root", root, "--overwrite")[0] == 0


def test_incompatible_output_root(tmp_path, capsys):
    (tmp_path / "f").write_text("x")
    code, _, err = run(capsys, "generate-sample", "--seed", "1", "--output-root", str(tmp_path / "f"))
    assert code == 2 and "no es un directorio" in err


def test_publication_cases_have_distinct_messages(tmp_path, monkeypatch, capsys):
    root = tmp_path / "d"
    real_replace, real_rmdir = syn._fs_replace, syn._fs_rmdir
    # caso 1: publicación interrumpida
    monkeypatch.setattr(syn, "_fs_replace", lambda s, d: (_ for _ in ()).throw(OSError(errno.EXDEV, "x")))
    code, _, err = run(capsys, "generate-sample", "--seed", "1", "--output-root", str(root))
    assert code == 2 and "manifest.json ausente" in err and "--overwrite" in err
    monkeypatch.setattr(syn, "_fs_replace", real_replace)
    # caso 3: conjunto publicado, limpieza fallida
    def boom(path):
        if Path(path).name == syn.STAGING_NAME:
            raise OSError(errno.EBUSY, "x")
        real_rmdir(path)
    monkeypatch.setattr(syn, "_fs_rmdir", boom)
    code, out, err = run(capsys, "generate-sample", "--seed", "1", "--output-root", str(root), "--overwrite")
    assert code == 2 and out == ""
    assert "publicado y validado" in err and "manifest.json presente" in err and "ausente" not in err
    assert (root / "ground_truth/manifest.json").exists() and (root / syn.STAGING_NAME).is_dir()
    monkeypatch.setattr(syn, "_fs_rmdir", real_rmdir)
    shutil.rmtree(root / syn.STAGING_NAME)
    assert run(capsys, "generate-sample", "--seed", "1", "--output-root", str(root), "--overwrite")[0] == 0
