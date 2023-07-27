Class Target
  Private m_name
  Public Sub Class_Initialize()
   debug.print "Target Init"
  End Sub
  Public Property Let name(in_name)
    debug.print "Target name set: " & in_name
    m_name = in_name
  End Property
  Public Property get name
    debug.print "Target name get: " & in_name
    out_name = m_name
  End Property
  Public Sub hit()
    debug.print "Target Hit: " & m_name
  End Sub
End Class
