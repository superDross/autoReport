from setuptools import setup, find_packages

setup(
    name='autoReport',
    version='0.1',
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        'Click',
        'openpyxl',
        'tqdm'
    ],
    entry_points='''
        [console_scripts]
        autoReport=autoReport.autoReport.autoReport:autoReport
    ''',
)
