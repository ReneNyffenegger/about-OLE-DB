option explicit

const HKEY_CLASSES_ROOT     = &H80000000

dim registry
set registry = getObject("winmgmts:\\.\root\default:StdRegProv")

dim allCLSIDs
registry.enumKey HKEY_CLASSES_ROOT, "CLSID", allCLSIDs

dim clsid
for each clsid in allCLSIDs

    dim dummy
    dim clsidPath

    clsidPath = "CLSID\" & clsid
    if registry.getDWordValue(HKEY_CLASSES_ROOT, clsidPath ,  "OLEDB_SERVICES",     dummy) = 0 or _
       registry.getDWordValue(HKEY_CLASSES_ROOT, clsidPath & "\OLEDB_SERVICES", "", dummy) = 0    _
    then

        dim providerName
        dim progId
        dim progIdNoVersion
        dim oleDbProvider

        registry.GetStringValue HKEY_CLASSES_ROOT, clsidPath                              , "", providerName  ' Default value
        registry.getStringValue HKEY_CLASSES_ROOT, clsidPath & "\OLE DB Provider"         , "", oleDbProvider ' Default value
        registry.getStringValue HKEY_CLASSES_ROOT, clsidPath & "\ProgID"                  , "", progId        ' Default value
        registry.GetStringValue HKEY_CLASSES_ROOT, clsidPath & "\VersionIndependentProgID", "", progIdNoVersion

        wscript.echo providerName
        wscript.echo "  progId  : " & progId  & " / " & progIdNoVersion
        wscript.echo "  Provider: " & oleDbProvider
        wscript.echo "  CLSID   : " & CLSID

        wscript.echo
    end if
next
