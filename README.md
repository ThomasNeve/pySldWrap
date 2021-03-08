# documentation

## Introduction
This project uses python to alter solidworks models through the solidworks API. The solidworks API is a COM interface to the solidworks software. The pywin32 project (win32com python library) is used to have python communicate to the COM interface. The python scripts can then be used to modify the parameters of the solidworks models.

## Folder structure
Project directory structure: the 'python_macros' directory contains the code from this repository while the 'slider_crank_models' folder should contain the solidworks models. The folder 'slider_crank_default' contains the assembly and parts that are used as a default model. The default model can then be modified and copied to the 'modified_models'.

```
project
|
|---python_macros
|   |
|   |   slider_crank.py
|   |   sw_tools.py
|   |   components.py
|   |   unit_test.py
|
|
|---slider_crank_models
    |
    |---modified_models
    |   |
    |   |   model_0
    |   |   model_1
    |   |   ...
    |
    |   slider_crank_default

```

## Installation
The necessary libaries can be found in the 'requirements.txt' file. This also requires installing the pywin32 project which can be done through pip. The installation steps are explained in [this](https://superuser.com/questions/609447/how-to-install-the-win32com-python-library) post and [here](https://github.com/mhammond/pywin32). This comes down to installing pywin32 through pip and running the post install.

```
1. python -m pip install pywin32

locate the Scripts folder from python (this can be the installation folder of python or the virtual environment folder):
2. python Scripts/pywin32_postinstall.py -install
```

The scripts were written for the API of solidworks 2019. If another version of solidworks is used, the version should be changed at the top of the sw_tools.py file.

The script should also be pointed to the correct slider_crank_default model folder. This can be done by setting the constant SLIDER_CRANK_DEFAULT_PATH in the *components*.py to the name of the folder that contains the default slider crank model.

## Geting started

Before running a script solidworks should be opened. This can be the default blank screen at startup.

![start_screen](./img/sw_starting_screen.PNG)

The python scripts are divided into several python files: *slider_crank*.py, *components*.py and *sw_tools*.py. The file *sw_tools*.py contains the functions that can be used to interact with the sw API, *slider_crank*.py contains the functions to modify the slider crank solidworks assembly and *components*.py contains the class that stores all the paths to the parts of the assembly that can be edited.

The slider crank can be modified by calling modify_assembly() with the desired parameters. If no path is passed then the default model is copied and modified.

```python
res = modify_assembly(dist_bottom_plate=0.45, crank_length=0.05, rod_length=0.3, path_model=None)
```

The result returns whether the modification was successful.

# solidworks-API (additional info)

The solidworks API is a COM interface to the solidworks software. The API only supports languages that support COM, e.g. VB.NET, visual c++ and C#. Here python was used with pywin32, a library that provides access to much of the win32 API and the ability to create and use COM objects. This, however, comes with some difficulties. Not all python bindings for the API work well all the time and might require converting some arguments before being passed to a method. This is also explained in detail in the following [post](http://joshuaredstone.blogspot.com/2015/02/solidworks-macros-via-python.html).


## win32com (python)

Might take a lot of time to write API wrapper to com interface ?

introduction to win32com with python: 

- http://timgolden.me.uk/pywin32-docs/html/com/win32com/HTML/QuickStartClientCom.html

- http://timgolden.me.uk/pywin32-docs/contents.html

additional info win32com (stackoverflow post):
- https://stackoverflow.com/questions/40660996/how-to-introspect-win32com-wrapper

- https://mail.python.org/pipermail/python-win32/2005-December/004031.html


## Example

```python
    import win32com.client
    import pythoncom
    swcom = win32com.client.Dispatch("SLDWORKS.Application");

    model = sw.ActiveDoc
    modelExt = model.Extension
    selMgr = model.SelectionManager
```

The model refers to the [IModelDoc2 Interface](http://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2.html) where the model instance refers to the active document in solidworks, i.e. the document that is currently opened in solidworks.

Members and functions from this interface can then be accessed through the model (com) object.

Example: call the [SelectByID2()](http://help.solidworks.com/2014/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imodeldocextension~selectbyid2.html) function from the [Extension](http://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2~Extension.html) interface to slecect a solidworks element through its ID.

```python
    arg1 = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    modelExt.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, False, 0, arg1, 0)
```

In this example the argument arg1 needs to be modified to a variant type, this is also explained in this [post](http://joshuaredstone.blogspot.com/2015/02/solidworks-macros-via-python.html). The arguments to the function need to be formed in the corrent type. This can be found in the makepy (.py) file from the win32com library. the necessary types can be found on the page [VARIANT Type Constants](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/3fe7db9f-5803-4dc4-9d14-5425d3f5461f?redirectedfrom=MSDN).


## check issues in model

An example can be found in this [post](https://forum.solidworks.com/thread/44251).
According to the API, a model can be rebuilt by the method [Rebuild()](https://help.solidworks.com/2019/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IModelDocExtension~Rebuild.html?verRedirect=1) and returns a bool value according to whether the rebuild is succesfull. This, however, does not seem to be the case in the python binding, a None value is returned. Issues can still be detected through the [GetWhatsWrongCount()](http://help.solidworks.com/2019/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDocExtension~GetWhatsWrongCount.html) method. There is also the [GetWhatsWrong()](http://help.solidworks.com/2019/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.imodeldocextension~getwhatswrong.html?verRedirect=1) method which returns the features that cause the issues.

Overlapping model parts are not detected as issues in the assembly. In order to prevent configurations where this is the case, limit mates can be introduced.

example: Use a limit mate to create sufficient distance between the slider support and the crank support.

![limit mate](./img/limit_mate.PNG)


## issues with packAndGo interface thourgh Com library

Can't acces packAndGo object through GetPackAndGo() as described in [this](https://forum.solidworks.com/thread/74577) post. 
Instead the assembly is copied by copying the full assmebly and parts directory to a new directory.

## Similarities between C# and VBA (NOT USED)

use Macro recorder and translate to C# -> call C# script from python ?

code examples:

- C#: https://help.solidworks.com/2020/English/api/sldworksapi/Get_Areas_of_MidSurface_Faces_Example_CSharp.htm

- VBA: https://help.solidworks.com/2020/English/api/sldworksapi/Get_Areas_of_MidSurface_Faces_Example_VB.htm


## solidworks modelling

changing the material of multiple parts macro: https://forum.solidworks.com/thread/26510

## usefull links

Dynamic dispatch and passing arguments by reference: https://github.com/mhammond/pywin32/issues/622