# pySldWrap

PySldWrap is a python library used for altering and interacting with SolidWorks models through the SolidWorks API. A number of python functions are implemented to interact with the SolidWorks software. This includes features like

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

## Getting started

Before running a script, SolidWorks should be opened. This can be the default blank screen at start up. An example on how to open and close a part is given below.

### opening/closing a part

```python
    import pySldWrap.sw_tools as sw_tools
    from pathlib import Path

    sw_tools.connect_sw("2019")  # open connection and pass Solidworks version

    path = 'part.SLDPRT'
    # path = Path(path)  # a path object can also be used for a number of functions
    model = sw_tools.open_part(path)  # open the model, link is returned
    sw_tools.close(path)  # close the model
```

### editing a part

A part can be modified when it is opened. When you are done editing a part, it should be saved before closing again.Saving can be done with save_model() or open_save_part(). The latter function also triggers a rebuild of the part before saving which could be necessary for some modifications.

```python
    path = 'part.SLDPRT'
    model = sw_tools.open_part(path)

    # the part can be edited here

    sw_tools.save_model(model)
    sw_tools.close(path)
```

Another convenient way of modifying a part is by using EditPart() which uses python's context manager.

```python
    path = 'part.SLDPRT'
    with sw_tools.EditPart(path) as model:
        # the part can be edited here
```

Upon entering the with block, the part is opened. Within this block the part can then be edited. The part is then automatically rebuild and saved before exiting the with block.

### modifying a sketch of a part

Lets say the part 'part.SLDPRT' has a sketch 'shape' with a dimension called 'length'. The value of this dimension can then be modified with the function edit_dimension_sketch().

```python
    new_length = 0.5
    sw_tools.edit_dimension_sketch(model, "shape", "length", new_length)
```

The part can then be rebuilt and saved with save_model().

```python
    sw_tools.open_save_part(model)
```

Or with the context manager.

```python
    path = 'part.SLDPRT'
    with sw_tools.EditPart(path) as model:
        new_length = 0.5
        sw_tools.edit_dimension_sketch(model, "shape", "length", new_length)
```

### modifying the value of an extrude

```python
    new_length = 0.35
    sw_tools.edit_dimension_extrude(model, "extrude_name", new_length)
```

### get the mass properties of a part

THe function mass_properties() extracts the mass properties along a certain coordinate system and returns the properties in a python dictionary. The properties COM, volume, surface, mass and moment of inertia I around all axes.

```python
    coord_sys_name = "CoordinateSystem_API"
    properties = sw_tools.mass_properties(model, coord_sys_name=coord_sys_name)
```

### exporting to .STEP

A part or assembly can be exported to a destination directory with export_to_step().

```python
    dst = './export_name.STEP'
    res = sw_tools.export_to_step(path_model, dst=dst)
```

### editing a pattern

```python
    sw_tools.edit_pattern(model, "pattern_name", D1TotalInstances=8)
```

### opening an assembly

```python
    path_asm = 'assembly.SLDASM'    # should be absolute path here
    sw_tools.open_assembly(path_asm)
```

More info on the available functions and their arguments can be found in the docstrings.

## How does it work

This library uses the pywin32 project (win32com python library) to communicate with the COM interface of the Solidworks API. Python functions are then wrapped around a subset of the Solidworks API.
