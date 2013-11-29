#!/usr/bin/env python3
# vim:fileencoding=utf-8
import sys

from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
buildOptions = dict(
  packages = [], excludes = [],
  include_files = ['resource', 'demo.xls', 'data.ini'],
)

name = 'main'

if sys.platform == 'win32':
  name = name + '.exe'

base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [
  Executable('main.py', base = base, targetName = name,
             compress = True, icon = "setup.ico",
            )
]

setup(name='main',
      version = '1.0',
      description = 'A little GUI program',
      options = dict(build_exe = buildOptions),
      executables = executables)
