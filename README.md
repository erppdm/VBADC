# VBADC
**V**isual **B**asic for **A**pplications **D**ynamic **C**ode

A tool to load VBA code dynamically at runtime.

## Why
I am facing a project in which about 80% of the programming will be implemented in VBA.
It is about hundreds of classes, modules, etc. The code is mainly used in Excel, Outlook, Word and SolidWorks-CAD.

On the one hand, this tool is used for simpler versioning of the source code at the touch of a button.
On the other hand, it is used to reload dynamically generated VBA source code at runtime.

## Prerequisites
- The macros **project name** has to be the same as the one in the variable **vbProject**.
- The filename of the INI file has to be the same as the filename in which the macro is stored.
- The INI file has to be located in the directory **ini**, which is located in the directory of the macro.
- Each module, class, etc. to be loaded is located after the key importMod[Counter] with the full filename in the section **Import**.
  ```
  [Import]
  importMod0="C:\git\VBA\ModOne.bas"
  importMod1="C:\git\VBA\ModTwo.bas"
  ```

## Limitations
Unfortunately, this tool cannot be used in Outlook.

To use this tool in Microsoft Office, **Trust access to the VBA project object model** has to be enabled.

Do these steps to allow it:
- Go to **File** > **Options** > **Trust Center**
- Click on **Trust Center Settings**
- Under **Macro Settings**, make sure **Trust access to the VBA project object model** is checked 

## Getting Started
- The files **BiIVbProjectManager.bas** and **BiIClassIni.cls** have to be imported into a VBA macro. 
- The **INI file** and the **variables** in the **module BiIVbProjectManager** has to be modified to your needs.
- Use **BiIVbProjectManager.ProjectLoader** to import
- Use **BiIVbProjectManager.ProjectExporter** to export
- That's all.
