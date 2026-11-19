def classFactory(iface):
    from .plugin import AutomationPackPlugin
    return AutomationPackPlugin(iface)
