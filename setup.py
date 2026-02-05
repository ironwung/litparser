"""
LitParser - Lightweight Document Parser
pip install -e . 또는 python setup.py install
"""

from setuptools import setup, find_packages
from pathlib import Path

readme = Path(__file__).parent / "README.md"
long_description = readme.read_text(encoding='utf-8') if readme.exists() else ""

setup(
    name="litparser",
    version="0.5.0",
    description="Lightweight Document Parser - 순수 Python으로 PDF, DOCX, PPTX, HWPX 파싱",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Your Name",
    author_email="your@email.com",
    url="https://github.com/yourusername/litparser",
    license="MIT",
    
    packages=find_packages(include=['litparser', 'litparser.*']),
    package_data={
        'litparser': ['*.md'],
    },
    
    python_requires=">=3.8",
    install_requires=[],
    
    extras_require={
        'dev': ['pytest>=7.0', 'pytest-cov>=4.0'],
    },
    
    entry_points={
        'console_scripts': [
            'litparser=litparser.__main__:main',
        ],
    },
    
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Topic :: Text Processing",
    ],
    
    keywords="pdf parser docx pptx hwpx document lightweight",
)
