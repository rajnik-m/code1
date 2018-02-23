Namespace Access.PostcodeValidation

  Friend Interface IPostcodeValidator
    Inherits IDisposable

    'Public Read/Write Properties
    Property CheckAddress As Boolean
    ReadOnly Property CanBulkAddressSearch As Boolean

    Function GetAddresses(ByVal pAddresses As IEnumerable(Of Interfaces.IAddress), ByRef qaVerifyLevel As Boolean) As IEnumerable(Of Interfaces.IAddress)
    Function PostcodeAddress(ByVal pAddress As Interfaces.IAddress) As IEnumerable(Of Interfaces.IAddress)
    Function ValidateBuilding(ByRef pAddress As PostcoderAddress) As Postcoder.ValidatePostcodeStatuses

  End Interface
End Namespace



