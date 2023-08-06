from setuptools import setup, find_packages

with open("requirements.txt", "r") as fh:
    requirements = fh.read().splitlines()

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
