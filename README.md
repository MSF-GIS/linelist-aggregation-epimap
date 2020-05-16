# Linelist Aggregation Epimap
> Data transformation and aggregation of Linelist for the Epimap

## Installation

I recommend you to install a Python environment with conda, virtualenv or pipenv.

### Conda
For example with conda, 
[download and install miniconda](https://docs.conda.io/en/latest/miniconda.html)

#### Option 1 : Using env.yml
Create the conda environment
```
conda create -n epimap -f env.yml
```

Activate the conda environment
```
activate epimap
```

#### Option 2 : Installing package by package
Create a conda environment
```
conda create -n epimap python=3.6.2
```

Activate the conda environment
```
activate epimap
```

Install dependencies
```
conda install pandas==0.20.3
conda install xlrd==1.2.0
conda install openpyxl==3.0.3
conda install python-levenshtein==0.12.0
```

### Virtualenv or pipenv

Setup a virtualenv or pipenv

Install dependencies
```
pip install -r requirements.txt
````

## Development Usage

Run the Python script aggregator.

```
python src/interfaceGraph.py 
```

To release, install pyinstaller `pip install pyinstaller`, and use the following command

````
pyinstaller src/interfaceGraph.py -F
````

## User usage

- Run `aggregation_epimap.exe` by double-clicking.
- You can find logs about all the executions in `logs.log` file.

## Tests

Go to the tests directory and run unittest
```
cd tests
python -m unittest discover
```

## License

The project has an [MIT license](Licence.md).
