# -*- coding: utf-8 -*-
import uno

def configuration_path(event=None):
    '''
    '''

    ctx = XSCRIPTCONTEXT.getComponentContext()
    doc = XSCRIPTCONTEXT.getDocument()

    # Get the configuration provider
    config_provider = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.configuration.ConfigurationProvider", ctx
    )

    # Access the PathSettings configuration
    path_settings = config_provider.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationAccess",
        (uno.createUnoStruct("com.sun.star.beans.NamedValue", "nodepath", "/org.openoffice.Office.Paths"),)
    )

    pip = ctx.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")

    text = doc.getText()  # com.sun.star.text.Text
    # Get the User-specific configuration path

    # text.setString(pip)
    # text.setString(pip.getPackageLocation("user"))

    for p in path_settings:
        text.setString("La sule les copains")
        text.setString("\n")


    # print("Configuration Path : ", conf_path)


# def apso_launcher(event=None):
#     '''
#     Allows to launch Apso from user interface (menu item, toolbar...).
#     '''
#     ctx = XSCRIPTCONTEXT.getComponentContext()
#     apso = ctx.ServiceManager.createInstance("apso.python.script.organizer.impl")
#     apso.trigger("execute")


g_exportedScripts = (configuration_path,)
