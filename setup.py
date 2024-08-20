from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {'packages': ['os','xlwings','datetime'], 'excludes': []}

base = 'gui'

executables = [
    Executable('main.py', base=base)
]

setup(name='Disney Offers Tool',
      version = '1.0',
      description = 'A Disney Offers tool used by Wideout/Media Ocean West Coast',
      options = {'build_exe': build_options},
      executables = executables)
