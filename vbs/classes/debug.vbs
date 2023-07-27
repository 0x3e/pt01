Class MyDebug
  Public Sub print(out)
    WScript.Echo out
  End Sub
  Public Sub wsh_info()
    print "debug world!"
    print "WSH Version: " &  WScript.Version
    print "ScriptEngine: " &  ScriptEngine & " " & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion
  End Sub
End Class

