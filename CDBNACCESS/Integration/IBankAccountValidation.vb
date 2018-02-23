Friend Interface IBankAccountValidation

  Sub ValidateBankAccount()

  'Public Read / Private Write Properties
  Property InvalidReasonCode As String
  Property InvalidReasonDesc As String
  Property InvalidParameterName As String
  Property SortCodeOutput As String
  Property AccountNumberOutput As String
  Property IBANOutput As String
  Property BankBICOutput As String
  Property BranchBICOutput As String
  Property BankName As String
  Property BranchTitle As String
  Property BranchName As String
  Property BranchAddressLine1 As String
  Property BranchAddressLine2 As String
  Property BranchAddressLine3 As String
  Property BranchAddressLine4 As String
  Property BranchTown As String
  Property BranchCounty As String
  Property BranchPostCode As String
  Property BranchCountryDesc As String
  Property BranchTelephone As String
  Property IsValid As Boolean
  Property VerifyResult As AccountNoVerify.AccountNoVerifyResult
  Property VerifyURL As String

  'Public properties
  Property SortCodeInput As String
  Property AccountNumberInput As String
  Property VerifyType As AccountNoVerify.UseVerifyType


End Interface
