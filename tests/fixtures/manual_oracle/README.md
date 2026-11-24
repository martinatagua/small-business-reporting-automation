# Fixture del oráculo manual (F3, DF3-10, control 3 de §10.3)

Escrito **a mano**: ningún archivo de esta carpeta lo genera `synthetic.py` ni ningún script. Contiene unos
16 registros que cubren R01–R14 (más F11) con su perfil propio y la tabla de resultados esperados, redactada
a mano a partir de §8.2 y §12.1. En F6 `test_manual_oracle` ejecutará el pipeline sobre `sales.csv` y comparará
con `expected.csv`; en F3 solo se comprueba su cobertura (T3-13).

- `config.toml`: perfil propio; mapea `reported_total` y `branch_in_file` y declara dos formatos de fecha.
- `catalog.csv`: catálogo mínimo (`AA-100`); `ZZ-999` no está en él (R12).
- `sales.csv`: fuente (UTF-8, `,`, decimal `.`). La fila 1 es la cabecera; `source_row` = número de registro.
- `expected.csv`: resultado esperado por registro (`level=record`) y por archivo (`level=file`).
