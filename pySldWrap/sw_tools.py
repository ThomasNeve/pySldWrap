import shutil
from pathlib import Path

import win32com.client
import pythoncom


class SW():
    
    def __init__(self) -> None:
        self.app = None
        
    def set_sw(self, sw):
        self.app = sw


sw = SW()

def connect_sw(sw_year):
    """connect to the solidworks API

    Args:
        sw_year (str): solidworks version (year), for example if you have solidworks 2019 pass "2019"
    """

    sw_app = win32com.client.Dispatch("SldWorks.Application.%d" % (int(sw_year)-2012+20))  # e.g. SW2012 is 20, SW2015 is 23

    sw.set_sw(sw_app)


class EditPart():
    """
    The class is used as a context manager to edit parts.
    The __enter__() method is called when the 'with' block is entered and the return value
    is passed to the variable after the 'as' keyword. When the block of code is executed or
    when an exception occurs, the __exit__() method is called. The return value determines whether
    to stop the exception or have it propagate further.
    """

    build_status = {}

    def __init__(self, path):
        self.path = path


    def __enter__(self):
        
        self.model = open_part(self.path)
        print('editing {}'.format(self.path.name))

        return self.model


    def __exit__(self, type, value, traceback):

        self.model.EditRebuild3

        EditPart.build_status[self.path] = True

        if (type is not None) or (value is not None) or (traceback is not None):
            EditPart.build_status[self.path] = False
            print('error occured while editing {}:'.format(self.path.name))
            print(value)

        if not save_model(self.model):
            EditPart.build_status[self.path] = False

        close(self.path.name)

        print()

        return True


def open_model(path):

    """
    Call open_part() or open_assembly() depending on wheter the file is a part or assembly.
    The model is not activated and displayed if the model was already open. However, a valid
    model pointer is still returned.

    Args:
        path (str): path to the model, can also be a Path object.

    Returns:
        The model pointer (IModelDoc2) if successful, None otherwise
    """

    if Path(path).suffix.upper() == '.SLDPRT':
        return open_part(path)
    else:
        return open_assembly(path)


def open_part(path):

    """
    Open the part at the given path.

    Args:
        path (str): the path to the part.
    """

    path = str((Path.cwd() / path).resolve())

    arg1 = win32com.client.VARIANT(pythoncom.VT_BSTR, path)
    arg2 = win32com.client.VARIANT(pythoncom.VT_I4, 1)
    arg3 = win32com.client.VARIANT(pythoncom.VT_I4, 1)
    arg5 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 2)
    arg6 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128)

    return sw.app.OpenDoc6(arg1, arg2, arg3, "", arg5, arg6)


def close(name):

    """
    Closes the open doc of the name that is given.

    Args:
        name (str): the filename of the part or assembly, can also be a Path object
    """

    if isinstance(name, Path):
        name = name.name

    sw.app.CloseDoc(name)


def open_assembly(abs_path):

    """
    Open the assembly at the given path.

    Args:
        path (str): absolute path to the assembly.
    """

    arg1 = win32com.client.VARIANT(pythoncom.VT_BSTR, abs_path)
    arg2 = win32com.client.VARIANT(pythoncom.VT_I4, 2)
    arg3 = win32com.client.VARIANT(pythoncom.VT_I4, 0) #1
    arg5 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 2)
    arg6 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128)

    return sw.app.OpenDoc6(arg1, arg2, arg3, "", arg5, arg6)


def activate_doc(name):

    """
    activate the doc of the name that is passed.
    The doc should already be opened.

    Args:
        name (str): The name of the doc, can be a str or the path object
    """

    if isinstance(name, Path):
        name = name.name

    arg1 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    return sw.app.ActivateDoc3(name, False, 2, arg1)


def save_model(model):

    """
    Save the model to the current file.

    Note:
        Saving an assembly will not save and rebuild all subassemblies and parts.
        Use rebuild_and_save_all() to rebuild and save all subassemblies and parts if necessary.

    Args:
        model (IModelDoc2): the model that is to be saved
    """

    arg1 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 1)
    arg2 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 1)
    arg3 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 1)
    model.save3(arg1, arg2, arg3)


def export_to_step(path_model, dst=Path('./model.STEP')):

    """
    Export the model, part or assembly, to a STEP format.

    Args:
        path_model (str): path to the model that is to be exported.
        dst (str, optional): path of the destination file with the filename and STEP extension,
            otherwise it is exported to the default location (./model.STEP).
    """

    path_model = Path(path_model)
    model = open_model(path_model)
    model = activate_doc(path_model) # activate the model if it was already opened

    extension = '.STEP'

    dst = Path.cwd() / dst
    if dst.suffix != extension:
        dst = dst.parent / (dst.name + extension)
    
    print('exporting to {}'.format(str(dst)))

    arg1 = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    arg2 = win32com.client.VARIANT(pythoncom.VT_BOOL, 0)
    arg3 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    arg4 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    ret = model.Extension.SaveAs2(str(dst), 0, 1, arg1, "", arg2, arg3, arg4)

    if not ret:
        print('export failed')
        return None

    return str(dst)


def open_save_part(path):

    """
    Open and save a part.

    Args:
        path (str): path to the model to be saved.
    """

    model = open_part(path)
    model.EditRebuild3
    save_model(model)

    close(path)


def open_save_assembly(path):

    """
    Open, rebuild and save the assembly at the given path.

    Args:
        path (str): path of the assembly file

    Returns:
        If there are errors or warnings in the build, False is returned,
        otherwise True is returned.
    """
    
    model = open_assembly(path)

    rebuild_status = model.EditRebuild3
    save_model(model)

    nr = model.Extension.GetWhatsWrongCount

    if nr>0:

        print('there are {} items with issues in the assembly'.format(nr))

        arg1 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0.0)
        arg2 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0.0)
        arg3 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0.0)

        if model.Extension.GetWhatsWrong(arg1, arg2, arg3):

            warnings = arg1.value
            err_code = arg2.value # error codes seem incorrect
            features = arg3.value

            feature_names = [feat.Name for feat in features]
            warnings = ['warning' if warning else 'error' for warning in warnings]

            problems = list(zip(warnings, err_code, feature_names))
            print('The following issues are present:', problems)


        # issues in the assembly
        return False


    # no issues in the assembly
    return True


def rebuild_and_save_all():

    """
    Iterate over all open documents and check if a model needs to rebuild and saved.
    The main assembly should first be rebuilt to detect what parts of the assembly need to be rebuilt and saved.
    """

    print('rebuilding and saving all necessary parts and assemblies')

    model = sw.app.GetFirstDocument

    while model is not None:
        
        path = Path(model.GetPathName)
        save_flag = model.GetSaveFlag

        if save_flag:

            print('rebuilding and saving:', str(path.resolve()))

            if (path.suffix).upper() == '.SLDPRT':
                open_save_part(path)
            else:
                open_save_assembly(path)


        model = model.GetNext


def edit_dimension_sketch(model, sketch, dim_id, val):

    """
    Edit the dimension of the sketch of the part and
    change the value of the dimension to the value that is passed.

    Args:
        model (IModelDoc2): pointer to the model of the sketch
        sketch (str): name of the sketch that is to be edited
        dim_id (str): the name of the dimension that needs to be changed
        val (float): new value of the dimension
    """

    arg1 = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    boolstatus = model.Extension.SelectByID2(sketch, "SKETCH", 0, 0, 0, False, 0, arg1, 0)

    feature = model.SelectionManager.GetSelectedObject6(1, -1)
    dim = feature.Parameter(dim_id)
    print('current value: {} m'.format(dim.SystemValue))

    errors = dim.SetSystemValue3(val, 1, None)

    model.EditRebuild3
    print('value is set to {} m'.format(dim.SystemValue))


def edit_dimension_extrude(model, extrude, val):

    """
    Edit the value of an extrude. This can be both a boss and cut extrude.

    Args:
        model (IModelDoc2): pointer to the model of the extrude.
        extrude (str): name of the extrude feature.
        val (float): new value of the extrude dimension.
    """

    arg1 = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    boolstatus = model.Extension.SelectByID2(extrude, "BODYFEATURE", 0, 0, 0, False, 0, arg1, 0)

    feature = model.SelectionManager.GetSelectedObject6(1, -1)
    feature_data = feature.getDefinition

    arg1 = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    is_good = feature_data.AccessSelections(model, arg1)

    forward = True
    depth = feature_data.getDepth(True)
    if not depth:
        forward = False # reverse direction
        depth = feature_data.getDepth(False)
    print('current value: {}'.format(depth))

    feature_data.SetDepth(forward, val)
    print('value is set to {}'.format(feature_data.GetDepth(forward)))

    arg1 = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    is_good = feature.ModifyDefinition(feature_data, model, arg1)

    feature_data.ReleaseSelectionAccess


def edit_pattern(model, pattern:str, **kwargs):
    """modify a linear pattern in an assembly

    Args:
        model (IModelDoc2): pointer to the model that contains the pattern
        pattern (str): the name of the pattern

        kwargs:
            D1ReverseDirection (bool): the direction from the selected edge
            D1Spacing (float): the spacing of the pattern
            D1TotalInstances (int): number of instances
            D2ReverseDirection (bool): the direction from the selected edge
            D2Spacing (float): the spacing of the pattern
            D2TotalInstances (int): number of instances
    
    Note:
        The pattern should be at the at the top level of the assembly,
        it cannot be in a subassembly.
        The function is currently only tested for the linear pattern,
        more info on patterns and their attributes:
        https://help.solidworks.com/2019/English/api/sldworksapiprogguide/Overview/Pattern_Features_and_their_Feature_Data_Objects.htm?id=3368f8e9d3374a6199746323ab9cf9b4
    """
    arg1 = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    boolstatus = model.Extension.SelectByID2(f"{pattern}", "COMPPATTERN", 0, 0, 0, False, 0, arg1, 0)
    
    feature = model.SelectionManager.GetSelectedObject6(1, -1)
    feature_data = feature.getDefinition

    # modify feature
    for key, value in kwargs.items():
        setattr(feature_data, key, value)

    arg1 = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    is_good = feature.ModifyDefinition(feature_data, model, arg1)


def mass_properties(model, coord_sys_name=None, intertia_com=False):

    """
    Return the mass properties for a given part. The properties are given with respect to a
    certain coordinate system as defined in the optional argument coord_sys_name.

    Args:
        model (IModelDoc2): pointer to the model.
        coord_sys_name (str, optional): name of the coordinate system around which
            the properties are calculated. By default around the origin.
        intertia_com (bool, optinal): The properties are defined around the center of mass
            if set to True, otherwise as defined in the option coord_sys_name.

    Returns:
        Dictionary with all the mass properties.
    """

    if not intertia_com:

        mass_property = model.Extension.CreateMassProperty

        if coord_sys_name:
            # change the default coordinate system
            coord_sys = model.Extension.GetCoordinateSystemTransformByName(coord_sys_name)
            mass_property.SetCoordinateSystem(coord_sys)

        com = mass_property.CenterOfMass
        comX = com[0]
        comY = com[1]
        comZ = com[2]

        V = mass_property.Volume
        surface = mass_property.SurfaceArea
        m = mass_property.Mass

        I = mass_property.GetMomentOfInertia(1)
        Ixx = I[0]
        Ixy = I[1]
        Ixz = I[2]
        Iyx = I[3]
        Iyy = I[4]
        Iyz = I[5]
        Izx = I[6]
        Izy = I[7]
        Izz = I[8]

        properties = {'comX':comX,
                    'comY':comY,
                    'comZ':comZ,
                    'V':V,
                    'surface':surface,
                    'm':m,
                    'Ixx':Ixx,
                    'Iyy':Iyy,
                    'Izz':Izz,
                    'Ixy':Ixy,
                    'Izx':Izx,
                    'Iyz':Iyz,
                    }

    else:

        arg1 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        properties = model.Extension.GetMassProperties2(1, arg1, False)

        comX = properties[0] # center of mass
        comY = properties[1]
        comZ = properties[2]
        V = properties[3]
        surface = properties[4]
        m = properties[5]
        Ixx = properties[6] # moments of inertia at the center of mass
        Iyy = properties[7]
        Izz = properties[8]
        Ixy = properties[9]
        Izx = properties[10]
        Iyz = properties[11]

        properties = {'comX':comX,
                    'comY':comY,
                    'comZ':comZ,
                    'V':V,
                    'surface':surface,
                    'm':m,
                    'Ixx':Ixx,
                    'Iyy':Iyy,
                    'Izz':Izz,
                    'Ixy':Ixy,
                    'Izx':Izx,
                    'Iyz':Iyz,
                    }

    return properties


def copy_assembly(src, dst):
    """
    Copy the directory passed to src to the dst directory and return the destination path.
    An exception is raised if the dst directory already exists.

    Args:
        src (str): The path of the directory that is copied.
        dst (str): The path of the destination directory.

    Returns:
        The path of the destination directory
    """
    dst = Path(dst)
    if dst.exists() and dst.is_dir():
        raise Exception('destination folder already exists')


    src = Path(src)
    shutil.copytree(src, dst)

    return dst


def replace_component(path_asm, part_id, replace_part_path, replace_all=False):

    """
    Replace the component, named part_id, of an assembly with a part at the path
    replace_part_path.

    Note:
        The component should be a top-level component. It cannot be a component of a sub-assembly.
        If a component of a sub-assembly needs to be replaced, open the sub-assembly instead and
        replace the component in that assembly. Afterwards the assembly should still be saved.

    Args:
        path_asm (str): path to the assembly to which the part belongs.
        part_id (str): name of the component in the assembly.
        replace_all (bool, optional): replace all instances of the selected component, default is False.

    Returns:
        bool: True if the replacement was successful.
    """
    
    asm = open_assembly(path_asm)

    components_asm = asm.GetComponents(True)
    components_names = [component.Name2 for component in components_asm]
    index_dash = [component.Name2.rfind('-') for component in components_asm]

    components_names_short = []
    for i in range(len(components_asm)):
        components_names_short.append(components_names[i][0:index_dash[i]])

    part_index = components_names_short.index(part_id)
    component = components_asm[part_index]

    arg1 = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    boolstatus = asm.Extension.SelectByID2(component.Name2, "COMPONENT", 0, 0, 0, False, 0, arg1, 0)

    arg1 = win32com.client.VARIANT(pythoncom.VT_BOOL, replace_all)
    arg2 = win32com.client.VARIANT(pythoncom.VT_I4, 0)
    arg3 = win32com.client.VARIANT(pythoncom.VT_BOOL, True)

    print('replacing with {}'.format(str(replace_part_path)))
    res = asm.ReplaceComponents2(str(replace_part_path), "", arg1, arg2, arg3)
    
    return res


def generatePartsList(path_asm):

    parts_list = []

    def returnParts(comp):

        components = comp.GetChildren
        if components is None:
            print('hello')

        if len(components):
            return components
        else:
            return []

            
    asm = open_assembly(path_asm)

    components_asm = list(asm.GetComponents(True))

    while len(components_asm):

        comp = components_asm.pop(0)
        components = returnParts(comp)

        if len(components):
            for comp in components:
                components_asm.append(comp)
        else:
            parts_list.append(comp.Name2)

    close(path_asm)

    print(f'There are {len(parts_list)} parts in the assembly')
    print(parts_list)


if __name__ == '__main__':

    pass
