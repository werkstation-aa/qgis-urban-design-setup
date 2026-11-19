def classFactory(iface):
    from .plugin import VectorPivotPlugin
    return VectorPivotPlugin(iface)
