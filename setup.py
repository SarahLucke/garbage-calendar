#!/usr/bin/env python3

from setuptools import setup

setup(
    name="GarbageCalendar",
    version="0.1",
    description="Tool to parse garbage collection schedule in ingolstadt into excel calendar",
    author="Sarah Lucke",
    author_email="sarah.lucke@googlemail.com",
    license="GNU GPL3",
    packages=["GarbageCalendar"],
    python_requires=">=3.6",
    install_requires=["openpyxl", "requests"],
    zip_safe=True,
)
