import subprocess
import os

# List of required packages with versions from pyproject.toml
required_packages = [
    "gitpython>=3.1.44",
    "openpyxl==3.1.5",
    "pandas==2.2.3",
    "pillow<11.0.0",
    "streamlit==1.31.0"
]

# Write the requirements
with open('requirements.txt', 'w') as f:
    for package in required_packages:
        f.write(f"{package}\n")

print("requirements.txt has been generated successfully!")
