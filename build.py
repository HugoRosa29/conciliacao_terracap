"""
Prepara a pasta docs/ publicada no GitHub Pages.

O núcleo em Python fica na raiz do projeto (fonte única da
verdade) e é copiado para docs/py/, que é a pasta que o
navegador baixa. Rode este script depois de alterar qualquer
módulo do núcleo:

    python build.py

Para checar, sem copiar, se as cópias estão em dia (útil em CI):

    python build.py --verificar
"""

from pathlib import Path
import shutil
import sys


RAIZ = Path(__file__).resolve().parent
DESTINO = RAIZ / "docs" / "py"

# Módulos que o navegador precisa. app.py (servidor) e
# validar.py (testes) ficam fora de propósito.
MODULOS = (
    "planilhas.py",
    "conciliacao.py",
    "exportar.py",
    "ponte.py",
)


def verificar():
    """
    Diz se docs/py/ está igual aos módulos da raiz.
    """

    desatualizados = []

    for modulo in MODULOS:

        origem = RAIZ / modulo
        copia = DESTINO / modulo

        if not copia.exists():
            desatualizados.append(f"{modulo} (faltando em docs/py/)")

        elif copia.read_bytes() != origem.read_bytes():
            desatualizados.append(f"{modulo} (diferente da raiz)")

    if desatualizados:
        print("docs/py/ está desatualizado:")

        for item in desatualizados:
            print("   -", item)

        print("\nRode: python build.py")

        return 1

    print(f"docs/py/ está em dia ({len(MODULOS)} módulos).")

    return 0


def copiar():
    """
    Copia os módulos do núcleo para docs/py/.
    """

    DESTINO.mkdir(parents=True, exist_ok=True)

    for modulo in MODULOS:

        origem = RAIZ / modulo

        if not origem.exists():
            print(f"ERRO: {modulo} não existe na raiz do projeto.")
            return 1

        shutil.copy2(origem, DESTINO / modulo)

        print(f"  copiado  {modulo}")

    # Impede o Jekyll do GitHub Pages de ignorar pastas e
    # arquivos que começam com _ ou .
    (RAIZ / "docs" / ".nojekyll").touch()

    print(f"\ndocs/py/ atualizado com {len(MODULOS)} módulos.")

    return 0


if __name__ == "__main__":
    if "--verificar" in sys.argv:
        sys.exit(verificar())

    sys.exit(copiar())
