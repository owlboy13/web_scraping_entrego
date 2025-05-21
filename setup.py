# SCRIPT PARA CONVERTER SCRIPT EM PYTHON EM CÓDIGO C


from setuptools import setup
from Cython.Build import cythonize

setup(
    ext_modules=cythonize("inativos_offs.py")
)

# comando - python setup.py build_ext --inplace