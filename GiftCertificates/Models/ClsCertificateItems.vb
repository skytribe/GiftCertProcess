Public Class ClsCertificateItems
    Public Property GCOrderId As Integer 'Gift Certificate Order Number eg.  00001
    Public Property ItemId As Integer    'Item ID from the Web Store eg. TDM10K, TDM12K, VID
    Public Property JRCertificateNumber As Integer   '//This is the newly assigned JUmprRun gift certificate ID assigned
    Public Property CertificateOrderReference As String '//Relates to GCOrderNumber+LineItemIndex eg.  00001-01,00001-2 etc.
    Public Property JumpRunCustomerID As Integer   '//Customer Id in JumpRun
    Public Property JumpRunItemId As Integer     '//ItemId in Jumprun to associate type of Gift Certificate or Discount Item
    Public Property Amount As Double
    Public Property DiscountAmount As Double     '//For Info Purpose only on Gift Certificate records
    Public Property DiscountCode As String       '// Web Store Discount Code eg. SANTA20, ELF20, CC20

    Public Property printCount As Integer = 0  '//TO IMPLEMENT 

End Class
