Class Targets
  Private m_targets
  Public Sub Class_Initialize()
    Set m_targets = CreateObject("Scripting.Dictionary")
    debug.print "Targets init"
  End Sub
  Public Sub create(name)
    debug.print "create: " & name
    Dim tmp_tar : Set tmp_tar = New Target
    tmp_tar.name = "test"
    m_targets.Add name, tmp_tar
  End Sub
  Public Sub hit(name)
    If m_targets.Exists(name) Then
      debug.print "Trying Hit: " & name
      m_targets.Item(name).hit()
    End If
  End Sub
End Class
