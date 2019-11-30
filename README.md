# Work-3
## Installation [on Windows PC]
1. [Anaconda](https://www.anaconda.com/distribution/#download-section)
> NOTE: In the last prompt, Tick the checkbox corresponding to "Add to path". This will enable using `conda` in the terminal, for installing/upgrading packages.

2. Packages required here. Just press the [install.bat](./installation/install.bat)
	- anytree
	- pywin32
	- xlwings
	- pandas

<div style="page-break-after: always;"></div>

## Execution
### M-1: via Excel RUN button [RECOMMENDED]
* <kbd>RUN</kbd> button is linked with `run.bat` file present in this directory, because the `octopus.dot` & `octopus.png` files were not created on execution.

### M-2: Build python code <kbd>ctrl + b</kbd>
This is creating the 2 files - `octopus.dot` & `octopus.png` as output. But it is not as interactive as M-1.


## Explain code
### 1. Python
### 2. Batch
```bat
@echo off
del /f octopus.dot octopus.png
python octopus.py
rem dot octopus.dot -T png -o octopus.png
rem pause
```

### 3. VBA
```vba
Sub RUN_Click()
    Shell "cmd.exe /k cd /d" & ThisWorkbook.Path & "&& run.bat" & "&& octopus.png && exit" & "&& exit"
End Sub
```


<!-- Reference 
https://stackoverflow.com/questions/51447235/python-not-able-to-graph-trees-using-graphviz-with-the-anytree-package -->