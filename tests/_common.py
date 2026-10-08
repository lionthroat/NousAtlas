"""Shared test helpers. Tests only ever touch COPIES in a temp folder."""
import os, shutil, sys, tempfile
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
REAL = r"C:\Users\wrong\Documents\The Understory (Obsidian Vault)\Biomes\Design Documents\Places\Redsaint Desert\Redsaint_Desert_Map_1.xlsx"
FIXTURE = os.path.join(ROOT, "tests", "fixtures", "redsaint.xlsx")


def sample_copy():
    """A throwaway copy of the sample workbook."""
    here = tempfile.mkdtemp(prefix="nousatlas-test-")
    src = FIXTURE if os.path.exists(FIXTURE) else REAL
    dst = os.path.join(here, "sample.xlsx")
    shutil.copyfile(src, dst)
    return dst
