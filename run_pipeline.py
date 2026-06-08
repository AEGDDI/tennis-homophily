import nbformat
from nbclient import NotebookClient
import os

def run_nb(rel_path, cwd=None):
    nb_path = os.path.abspath(rel_path)
    if cwd is None:
        cwd = os.path.dirname(nb_path)
    with open(nb_path, encoding="utf-8") as f:
        nb = nbformat.read(f, as_version=4)
    client = NotebookClient(nb, timeout=300, kernel_name="python3",
                            resources={"metadata": {"path": cwd}})
    client.execute()
    with open(nb_path, "w", encoding="utf-8") as f:
        nbformat.write(nb, f)
    print(f"Done: {rel_path}")

root = os.getcwd()

run_nb("code/merging/merge_matches.ipynb",  cwd=os.path.join(root, "code", "merging"))
run_nb("code/cleaning/matches_doubles.ipynb", cwd=os.path.join(root, "code", "cleaning"))
run_nb("code/merging/matches_ranks.ipynb",  cwd=os.path.join(root, "code", "merging"))
run_nb("code/cleaning/final_ds.ipynb",       cwd=os.path.join(root, "code", "cleaning"))
run_nb("code/cleaning/tiebreak_panel.ipynb", cwd=os.path.join(root, "code", "cleaning"))
run_nb("code/analysis/homophily.ipynb", cwd=os.path.join(root, "code", "analysis"))
