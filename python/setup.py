from setuptools import setup, find_packages
from cffi import FFI

ffi = FFI()

with open('cdefs/spreadsheet_2007/Excel.cdef', 'r') as f:
    ffi.cdef(f.read())

ffi.set_source(
    "openxmloffice_c_ffi",
    """
    // No header file needed since the function signatures are defined in cdef.
    """,
    libraries=["openxmloffice_ffi"],
    library_dirs=["lib"]
)

setup(
    name="draviavemal_openxml_office",
    version="4.0.0",
    author="Dravia Vemal",
    author_email="contact@draviavemal.com",
    description="A Python wrapper for a openxml-office rust library.",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    install_requires=["cffi>=1.17.0"],
    setup_requires=["cffi>=1.17.0"],
    packages=find_packages(),
    ext_modules=[ffi.distutils_extension()],
    include_package_data=True,
    classifiers=[
        "openxml-office",
    ],
    python_requires=">=3.6",
)
