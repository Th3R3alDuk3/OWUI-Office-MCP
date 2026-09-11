from os import environ, execv
from pathlib import Path
from sys import argv

from py_landlock import Landlock


# This interpreter has no worker threads; OfficeCLI replaces it after confinement.
(
    Landlock()
    # Screenshots need the font configuration; without it every font falls back.
    .allow_read("/usr", "/proc", "/dev", "/etc/fonts")
    .allow_execute("/usr")
    .allow_read_write(Path.cwd(), environ["TMPDIR"])
    .apply()
)

execv("/usr/local/bin/officecli", ["officecli", *argv[1:]])
