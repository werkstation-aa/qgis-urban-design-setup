def classFactory(iface):
    from .plugin import LupTablesPlugin
    return LupTablesPlugin(iface)
