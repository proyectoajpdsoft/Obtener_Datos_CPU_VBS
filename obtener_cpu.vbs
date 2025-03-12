' Obtenemos mediante WMI los datos de los procesadores (CPU) del equipo actual
equipo = "."
set obCPU = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & equipo & "\root\cimv2")
set lsCPU = obCPU.ExecQuery("Select name, maxclockspeed, caption, family, description, deviceid, manufacturer, processortype, socketdesignation, numberofcores, numberoflogicalprocessors from Win32_Processor")

' Comprobamos que no se haya producido error al obtener los datos por WMI
on error resume next
siError = lsCPU.Count
if (err.number <> 0) then
  siError = true
else
  siError = false
end if
on error goto 0 

if (not siError) then
	' Mostramos los datos de todos los procesadores
	for each cpu in lsCPU
	  Wscript.StdOut.WriteLine "Procesador ID: " & cpu.deviceid
	  Wscript.StdOut.WriteLine "  -> Nombre: " & cpu.name
	  Wscript.StdOut.WriteLine "  -> Velocidad: " & cpu.maxclockspeed & " MHz"
	  Wscript.StdOut.WriteLine "  -> Familia: " & cpu.family
	  Wscript.StdOut.WriteLine "  -> Fabricante: " & cpu.manufacturer
	  Wscript.StdOut.WriteLine "  -> Tipo: " & cpu.processortype
	  Wscript.StdOut.WriteLine "  -> Socket: " & cpu.socketdesignation
	  Wscript.StdOut.WriteLine "  -> Número de cores: " & cpu.numberofcores
	  Wscript.StdOut.WriteLine "  -> Número de procesadores lógicos: " & cpu.numberoflogicalprocessors
	next
end if