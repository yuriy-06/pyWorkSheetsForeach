from setuptools import setup, find_packages
from os.path import join, dirname
import WorkSheetsForeach
setup(
    name='WorkSheetsForeach',
    version=WorkSheetsForeach.__version__,
    packages=find_packages(),
    long_description=open(join(dirname(__file__), 'README.txt')).read(),
	install_requires=[
    'pypiwin32']
)