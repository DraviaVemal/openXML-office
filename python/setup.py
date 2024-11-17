from setuptools import setup, find_packages

setup(
    name="draviavemal_openxml_office",
    version="4.0.0",
    author="Dravia Vemal",
    author_email="contact@draviavemal.com",
    description="A Python wrapper for a openxml-office rust library.",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    packages=find_packages(),
    include_package_data=True,
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.6",
)
