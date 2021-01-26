 $driveEject = New-Object -ComObject Shell.Application

 $driveEject.NameSpace(17).ParseName("D:").InvokeVerb("Eject")