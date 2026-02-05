from setuptools import setup, find_packages
from pathlib import Path

readme = Path(__file__).parent / "README.md"
long_description = readme.read_text(encoding='utf-8') if readme.exists() else ""

setup(
    name="litparser",
    version="0.5.3",
    description="Lightweight Document Parser - 순수 Python으로 PDF, DOCX, PPTX, HWPX 파싱",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="ironwung",
    author_email="ironwung@gmail.com",
    url="https://github.com/ironwung/litparser",
    license="MIT",
    
    packages=find_packages(),
    
    python_requires=">=3.8",
    install_requires=[],
    
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
