from distutils.core import setup
import py2exe, sys, os, glob
sys.argv.append('py2exe')



setup(
    options = {'py2exe': {'bundle_files': 1, 'compressed': True}},
    windows = [{"script" : "Priority_report.py"}],
    zipfile = None,
    version='1.0.0',
)

# setup(console=['Priority_report.py'])
