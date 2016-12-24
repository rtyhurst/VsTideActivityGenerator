try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

config = {
    'description': 'Tide Newsletter Activity generator reads a cvs file of '
                   'Real Estate Actives & Sold properties to generate a '
                   'Activity report for inclusion in a Tide Newsletter',
    'author': 'Rick Tyhurst',
    'url': 'URL to get it at.',
    'download_url': 'Where to download it.',
    'author_email': 'rtyhurst@gmail.com',
    'version': '0.1',
    'install_requires': ['nose','openpyxl'],
    'packages': ['NAME'],
    'scripts': [],
    'name': 'tideActivityGenerator'
}

setup(**config)
