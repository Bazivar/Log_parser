from cx_Freeze import setup, Executable

executables = [Executable('parser.py',targetName='Parser.exe', base='Win32GUI', icon='Unibelus_ico.ico')]



excludes = ['unicodedata', 'logging', 'unittest', 'email', 'html', 'http', 'urllib',
            'xml', 'pydoc', 'doctest', 'argparse', 'datetime', 'zipfile',
            'subprocess', 'pickle', 'threading', 'locale', 'calendar', 'functools',
            'weakref', 'tokenize', 'base64', 'gettext', 'heapq', 're', 'operator',
            'bz2', 'fnmatch', 'getopt', 'reprlib', 'string', 'stringprep',
            'contextlib', 'quopri', 'copy', 'imp', 'keyword', 'linecache']

options = {
    'build_exe': {
        'include_msvcr': True
    }
}

setup(name='Парсер логов',
      version='0.0.1',
      description='Парсер логов программы',
      executables=executables,
      options=options)