# pySldWrap

Python library to alter solidworks models through the solidworks API which works via the windows COM interface. The pywin32 project (win32com python library) is used to have python communicate to the COM interface. The python scripts can then be used to modify the parameters of the solidworks models. A number of functions are implemented to interact with the Solidworks software. This includes features like

- opening parts and assemblies
- modifying sketches and extrudes
- exporting to .STEP
- exporting the mass properties of a part
- replacing parts of an assembly

## Installation

The package can be installed through pip by running the following command.

```sh
    pip install pySldWrap
```

## Geting started

Before running a script Solidworks should be opened. This can be the default blank screen at start up. An example on how to open and close a part is given below.

```python
    import pySldWrap.sw_tools as sw_tools
    from pathlib import Path

    sw_tools.connect_sw("2019")  # open connection and pass Solidworks version

    path = 'part.SLDPRT'
    # path = Path(path)  # a path object can also be used for a number of functions
    model = sw_tools.open_part(path)  # open the model, link is returned
    sw_tools.close(path)  # close the model
```

More info on the functions can be found in the docstrings.