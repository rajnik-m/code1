Namespace Config
  ''' <summary>
  ''' This enum represents the different values that a Config Name's Config Scope can have
  ''' </summary>
  ''' <remarks>This is a flags enum.  
  ''' As configs are always available to the system, its flag is 0.
  ''' The actual value set against the config comes from a maintenance lookup, which should never be the exact values below, 
  ''' but a bitwise combination of them.
  ''' Expected stored values are
  ''' </remarks>
  <Flags>
  Public Enum ConfigNameScope
    SystemOnly = 0 'Can be stored against the config.
    Department = 1 'Can be stored against the config.As System is 0, Department actually means System And Department.
    User = 2 'This flag should never be stored against the config name.  It's here in order to make the bitwise calculations work
    SystemAndDepartmentAndUser = 3 'Can be stored against the config.
  End Enum
End Namespace