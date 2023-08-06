from setuptools import setup, find_packages

with open("requirements.txt", "r") as fh:
    requirements = fh.read().splitlines()

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="pyvfp",
    version="0.1.0",
    packages=find_packages(),
    package_data={
        'pyvfp': ['bin/*'],
    },
    install_requires=requirements,
    include_package_data=True
)
