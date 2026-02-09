def classFactory(iface):
    from .main_plugin import CadastralAuditor
    return CadastralAuditor(iface)