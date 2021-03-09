# pySldWrap

Python library to alter solidworks models through the solidworks API which works via the windows COM interface. The pywin32 project (win32com python library) is used to have python communicate to the COM interface. The python scripts can then be used to modify the parameters of the solidworks models.

## Installation

```sh
    pip install pySldWrap
```

## Geting started

Before running a script solidworks should be opened. This can be the default blank screen at startup.

## Example

```python
    import pySldWrap.sw_tools
```
