from setuptools import setup

setup(
    name='convert_to_audio',
    version='1.0',
    py_modules=['convert_to_audio'],
    install_requires=[
        'pyttsx3',
        'python-docx',
    ],
    entry_points={
        'console_scripts': [
            'convert_to_audio = convert_to_audio:main',
        ],
    },
)
