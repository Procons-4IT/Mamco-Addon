'//  SAP BUSINESS ONE MAMCO ADDONS
'//****************************************************************************
'//
'//  File:      MamcoAddOns.vb
'//
'//****************************************************************************
'// BEFORE STARTING:
'// 1. Add reference to the "SAP Business One UI API"
'// 2. Insert the development connection string to the "Command line argument"
'//-----------------------------------------------------------------
'// 1.
'//    a. Project->Add Reference...
'//    b. select the "SAP Business One UI API 6.7" From the COM folder
'//
'// 2.
'//     a. Project->Properties...
'//     b. choose Configuration Properties folder (place the arrow on Debugging)
'//     c. place the following connection string in the 'Command line arguments' field
'// 0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056
'//
'//**************************************************************************************************

Public Class MamcoAddOns

    Private WithEvents SBOApplication As SAPbouiCOM.Application ' UI application
    Private SBOCompany As SAPbobsCOM.Company      ' DI application

    Private Const frmBatchSelection As String = "42"
    Private Const frmPickManager As String = "81"
    Private Const frmPickList As String = "85"
    Private Const frmARInvoice As String = "133"
    Private Const frmBusinessPartner As String = "134"
    Private Const frmDelivery As String = "140"
    Private Const frmPurchaseOrders As String = "142"
    Private Const frmGoodsReceiptPO As String = "143"
    Private Const frmItemMaster As String = "150"
    Private Const frmOutgoingPayments As String = "169"
    Private Const frmIncomingPayments As String = "170"
    Private Const frmARCreditMemo As String = "179"
    Private Const frmInventoryTransfer As String = "940"
    Private Const frmDocumentDrafts As String = "3002"
    Private Const frmPickDetails As String = "60020"
    Private Const frmARInvoicePayment As String = "60090"
    Private Const frmARReverseInvoice As String = "60091"
    Private Const frmLandedCosts As String = "11032"
    Private Const frmCFLTerritories As String = "10199"
    Private Const frmCFLSuppliers As String = "xxxxx"
    Private Const frmShipment As String = "MAM_FSHP"


    Private Const ctlVenCode As String = "101"
    Private Const ctlVenName As String = "111"
    Private Const ctlCnctCode As String = "121"
    Private Const ctlMatrix As String = "301"

    Private Const btnCopyItems As String = "3"
    Private Const btnDraft As String = "btnDraft"

    Private Const lnkCnctCode As String = "122"

    Private Const mnuPurchasing As String = "2304"
    Private Const mnuAddRecord As String = "1282"
    Private Const mnuShipment As String = "MAM_MSHP"

    Private Const selVendor As String = "2"

    '**********************************************************
    ' declaring an Event filters container object and an
    ' event filter object
    '**********************************************************

    Public oFilters As SAPbouiCOM.EventFilters
    Public oFilter As SAPbouiCOM.EventFilter

    Private oCFLForm As SAPbouiCOM.Form
    Private oShipmentForm As SAPbouiCOM.Form    ' The shipment form
    Private oMatrix As SAPbouiCOM.Matrix        ' Global variable to handle matrixes
    Private oVenCode As SAPbouiCOM.EditText    ' Global variable for the BP Code
    ' Variables for Mamco Add Ons
    Private cmbBPCode As SAPbouiCOM.ComboBox    ' Global variable for the BP Code
    Private oDraftBtn As SAPbouiCOM.Button    ' Global variable for the BP Code
    Private txtBPName As SAPbouiCOM.EditText    ' Global variable for the BP Code
    Private colItemCode As SAPbouiCOM.Column    ' Global variable for the Item Code
    Private colItemName As SAPbouiCOM.Column    ' Global variable for the Item Name
    Private colItemPrice As SAPbouiCOM.Column   ' Global variable for the Item Price
    Private colItemQuan As SAPbouiCOM.Column    ' Global variable for the Item Quantity
    Private colInitQuan As SAPbouiCOM.Column    ' Global variable for the Item Initial Quantity
    Private colItemTotal As SAPbouiCOM.Column   ' Global variable for the Item Total
    Private oDocTotal As SAPbouiCOM.EditText    ' Global variable for the Blanket Total
    Private oCFLs As SAPbouiCOM.ChooseFromListCollection ' Global variable for form choose from lists

    Private nDraftKey As Long                    ' Draft Document saving document is based
    Private AddStarted As Boolean                ' Flag that indicates "Add" process started
    Private RedFlag As Boolean                   ' RedFlag when true indicates an error during "Add" process
    Private InvOk As Boolean                     ' Indicates that an invoice can be withdrawn from blanket agreement
    Private RowNum As Integer                    ' Row Number in a matrix
    Private OrderCodes() As String               ' Array to render Order Codes
    Private ItemCodes() As String                ' Array to render Item Codes
    Private ItemQuantities() As Double           ' Array to render Item Quantities
    Private oEditBPCode As SAPbouiCOM.EditText   ' Global variable for the BP code edit text
    Private oEditPostDate As SAPbouiCOM.EditText ' Global variable for the Post Date edit text
    Private strEditBPCode As String              ' Global variable for the BP code edit text string value
    Private strEditPostDate As String            ' Global variable for the Post Date edit text string value
    Private IsNewItem As Boolean                 ' Indicates if you added a new item to invoice/quotation/order
    Private PickQtyChanged As Boolean                ' Flag that indicates "Add" process started
    Private isBatchUpdatedAutomatically As Boolean                ' Flag that indicates "Add" process started
    Private isARInvoicePaymentLoading As Boolean                ' Flag that indicates "Add" process started

    Private NumberProvider As System.Globalization.NumberFormatInfo

#Region "Single Sign On"
    Private Sub SetApplication()
        AddStarted = False
        RedFlag = False

        '*******************************************************************
        '// Use an SboGuiApi object to establish connection
        '// with the SAP Business One application and return an
        '// initialized appliction object
        '*******************************************************************
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String
        SboGuiApi = New SAPbouiCOM.SboGuiApi
        '// by following the steps specified above, the following
        '// statment should be suficient for either development or run mode
        sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
        '// connect to a running SBO Application
        SboGuiApi.Connect(sConnectionString)
        '// get an initialized application object
        SBOApplication = SboGuiApi.GetApplication()
    End Sub

    Private Function SetConnectionContext() As Integer
        Dim sCookie As String
        Dim sConnectionContext As String
        Dim lRetCode As Integer
        Try
            '// First initialize the Company object
            SBOCompany = New SAPbobsCOM.Company
            '// Acquire the connection context cookie from the DI API.
            sCookie = SBOCompany.GetContextCookie
            '// Retrieve the connection context string from the UI API using the
            '// acquired cookie.
            sConnectionContext = SBOApplication.Company.GetConnectionContext(sCookie)
            '// before setting the SBO Login Context make sure the company is not
            '// connected
            If SBOCompany.Connected = True Then
                SBOCompany.Disconnect()
            End If
            '// Set the connection context information to the DI API.
            SetConnectionContext = SBOCompany.SetSboLoginContext(sConnectionContext)

        Catch ex As Exception
            SBOApplication.StatusBar.SetText("Mamco Addons failed setting a connection to DI API: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            SetConnectionContext = -1
        End Try
    End Function

    Private Function ConnectToCompany() As Integer
        Dim strCookie As String
        Dim strConnectionContext As String

        Try
            ConnectToCompany = SBOCompany.Connect
        Catch ex As Exception
            ConnectToCompany = -1
        End Try
        Try
            For intRow As Integer = 0 To 100
                If ConnectToCompany <> 0 Then
                    Try
                        ConnectToCompany = SBOCompany.Connect
                    Catch ex As Exception
                        ConnectToCompany = -1
                    End Try
                Else
                    Exit For
                End If
            Next
        Catch ex As Exception
            ConnectToCompany = -1
        End Try

        Try
            If ConnectToCompany <> 0 Then
                ConnectToCompany = SBOCompany.Connect
            End If
        Catch ex As Exception
            SBOApplication.StatusBar.SetText("Mamco AddOns failed to connect: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            ConnectToCompany = -1
        End Try
        


    End Function

    Private Sub Class_Init()

        '//*************************************************************
        '// set SBOApplication with an initialized application object
        '//*************************************************************

        SetApplication()

        '//*************************************************************
        '// Set The Connection Context
        '//*************************************************************

        If Not SetConnectionContext() = 0 Then
            SBOApplication.MessageBox("Failed setting a connection to DI API")
            End ' Terminating the Add-On Application
        End If
        '//*************************************************************
        '// Connect To The Company Data Base
        '//*************************************************************
        If Not ConnectToCompany() = 0 Then
            SBOApplication.MessageBox("Failed connecting to the company's Data Base")
            End ' Terminating the Add-On Application
        End If

        Dim oCompanyService As SAPbobsCOM.CompanyService
        oCompanyService = SBOCompany.GetCompanyService()

        Dim oAdminInfo As SAPbobsCOM.AdminInfo
        oAdminInfo = oCompanyService.GetAdminInfo()

        'SBOFunctions.sDecimalSeparator = oAdminInfo.DecimalSeparator
        'SBOFunctions.sThousandsSeparator = oAdminInfo.ThousandsSeparator

        NumberProvider = New System.Globalization.NumberFormatInfo
        NumberProvider.NumberDecimalSeparator = oAdminInfo.DecimalSeparator
        NumberProvider.NumberGroupSeparator = oAdminInfo.ThousandsSeparator
        NumberProvider.NumberGroupSizes = New Integer() {3}

        '//*************************************************************
        '// send a connected message
        '//*************************************************************

        SBOApplication.StatusBar.SetText("Addons for " & SBOCompany.CompanyName & vbNewLine & " are loaded!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


    End Sub
#End Region

    Public Sub New()
        MyBase.New()
        Class_Init()
        'AddMenuItems()
        SetFilters()
    End Sub

    Private Sub SetFilters()

        oFilters = New SAPbouiCOM.EventFilters

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD)
        oFilter.AddEx(frmPurchaseOrders)
        oFilter.AddEx(frmGoodsReceiptPO)
        oFilter.AddEx(frmInventoryTransfer)
        oFilter.AddEx(frmPickList)
        oFilter.AddEx(frmARInvoice)
        oFilter.AddEx(frmARInvoicePayment)
        oFilter.AddEx(frmARCreditMemo)
        oFilter.AddEx(frmARReverseInvoice)
        oFilter.AddEx(frmBusinessPartner)
        'oFilter.AddEx(frmShipment)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
        oFilter.AddEx(frmPurchaseOrders)
        oFilter.AddEx(frmGoodsReceiptPO)
        oFilter.AddEx(frmInventoryTransfer)
        oFilter.AddEx(frmPickDetails)
        oFilter.AddEx(frmARInvoice)
        oFilter.AddEx(frmARInvoicePayment)
        oFilter.AddEx(frmARCreditMemo)
        oFilter.AddEx(frmARReverseInvoice)
        oFilter.AddEx(frmBusinessPartner)
        'oFilter.AddEx(frmShipment)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
        oFilter.AddEx(frmARInvoice)
        oFilter.AddEx(frmARInvoicePayment)
        oFilter.AddEx(frmARCreditMemo)
        oFilter.AddEx(frmARReverseInvoice)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
        oFilter.AddEx(frmPurchaseOrders)
        oFilter.AddEx(frmGoodsReceiptPO)
        oFilter.AddEx(frmInventoryTransfer)
        oFilter.AddEx(frmBatchSelection)
        oFilter.AddEx(frmItemMaster)
        oFilter.AddEx(frmPickList)
        oFilter.AddEx(frmPickDetails)
        oFilter.AddEx(frmARInvoice)
        oFilter.AddEx(frmARInvoicePayment)
        oFilter.AddEx(frmARCreditMemo)
        oFilter.AddEx(frmARReverseInvoice)
        oFilter.AddEx(frmBusinessPartner)
        oFilter.AddEx(frmDelivery)
        'oFilter.AddEx(frmLandedCosts)
        'oFilter.AddEx(frmOutgoingPayments)
        'oFilter.AddEx(frmIncomingPayments)
        'oFilter.AddEx(frmShipment)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
        oFilter.AddEx(frmCFLTerritories)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)
        oFilter.AddEx(frmARInvoicePayment)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        oFilter.AddEx(frmPurchaseOrders)
        oFilter.AddEx(frmGoodsReceiptPO)
        oFilter.AddEx(frmInventoryTransfer)
        oFilter.AddEx(frmBatchSelection)
        oFilter.AddEx(frmPickManager)
        oFilter.AddEx(frmARInvoice)
        oFilter.AddEx(frmARInvoicePayment)
        oFilter.AddEx(frmARCreditMemo)
        oFilter.AddEx(frmARReverseInvoice)
        oFilter.AddEx(frmDelivery)
        'oFilter.AddEx(frmShipment)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK)
        oFilter.AddEx(frmBatchSelection)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE)
        oFilter.AddEx(frmPickList)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
        oFilter.AddEx(frmARInvoice)
        oFilter.AddEx(frmARInvoicePayment)
        oFilter.AddEx(frmARCreditMemo)
        oFilter.AddEx(frmARReverseInvoice)
        oFilter.AddEx(frmPickList)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
        oFilter.AddEx(frmARInvoicePayment)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
        oFilter.AddEx(frmARInvoice)
        oFilter.AddEx(frmARInvoicePayment)
        oFilter.AddEx(frmARCreditMemo)
        oFilter.AddEx(frmARReverseInvoice)
        oFilter.AddEx(frmShipment)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK)
        'oFilter.AddEx(frmShipment)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
        oFilter.AddEx(mnuAddRecord)

        SBOApplication.SetFilter(oFilters)

    End Sub

    ' This procedure adds a new Blanket Agreement form to the SBO UI
    Private Sub DrawForm()
        Dim oItem As SAPbouiCOM.Item               ' An item on the new form
        Dim oCombo As SAPbouiCOM.ComboBox          ' Combo box item for the BP code and Item code
        Dim oEditText As SAPbouiCOM.EditText
        Dim oColumns As SAPbouiCOM.Columns         ' The Columns collection on the matrix
        Dim oStaticText As SAPbouiCOM.StaticText
        Dim oComboBox As SAPbouiCOM.ComboBox

        Try
            LoadFromXML("frmShipment.srf")
        Catch ex As Exception
            SBOApplication.MessageBox(ex.Message)
        End Try

        oShipmentForm = SBOApplication.Forms.Item(frmShipment)

        oCFLs = oShipmentForm.ChooseFromLists

        SetListCondition(oCFLs.Item("2"), "CardType", "S")
        SetListCondition(oCFLs.Item("4"), "CardType", "S")

        ' Add Edit Items

        ' BP Code
        'oItem = oForm.Items.Item("txtCode")
        'cmbBPCode = oItem.Specific
        'AddBPCodeCombo(cmbBPCode)

        ' BP Name
        'oItem = oForm.Items.Item("txtName")
        'txtBPName = oItem.Specific

        ' Doc Total
        'oItem = oForm.Items.Item("txtTotal")
        'oDocTotal = oItem.Specific

        ' Add a matrix
        oMatrix = oShipmentForm.Items.Item(ctlMatrix).Specific
        oColumns = oMatrix.Columns

        oVenCode = oShipmentForm.Items.Item(ctlVenCode).Specific

        'colItemCode = oColumns.Item("ItemCode")
        'colItemName = oColumns.Item("ItemName")
        'colItemPrice = oColumns.Item("ItemPrice")
        'colItemQuan = oColumns.Item("ItemQuan")
        'colInitQuan = oColumns.Item("InitQuan")
        'colItemTotal = oColumns.Item("ItemTotal")

        'Add Valid Values
        'AddItemsToCombo(colItemCode)
        'oMatrix.AddRow()

        oShipmentForm.PaneLevel = 1

    End Sub

    ' This procedure adds all the items codes and names
    Private Sub AddItemsToCombo(ByVal oColumn As SAPbouiCOM.Column)
        Dim RS As SAPbobsCOM.Recordset
        Dim Bob As SAPbobsCOM.SBObob

        RS = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Bob = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)

        RS = Bob.GetItemList

        RS.MoveFirst()
        While RS.EoF = False
            oColumn.ValidValues.Add(RS.Fields.Item("ItemCode").Value, RS.Fields.Item("ItemName").Value)
            RS.MoveNext()
        End While
    End Sub

    ' This procedure adds all the items codes and names
    Private Sub SetListCondition(ByRef oCFL As SAPbouiCOM.ChooseFromList, ByVal sAlias As String, ByVal sCondVal As String)
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition

        oCons = New SAPbouiCOM.Conditions

        oCon = oCons.Add()
        oCon.Alias = sAlias
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon.CondVal = sCondVal
        oCFL.SetConditions(oCons)

    End Sub

    Private Sub EnableCopyItemsButton()
        Dim oForm As SAPbouiCOM.Form
        oForm = SBOApplication.Forms.Item(frmShipment)
        Dim oItem As SAPbouiCOM.Item
        oItem = oForm.Items.Item(btnCopyItems)
        oItem.Enabled = True
    End Sub

    Private Sub UpdateMatrixColumns()

        Dim colMark As SAPbouiCOM.Column
        colMark = oMatrix.Columns.Item("0")
        Dim nRow As Integer
        Dim oEdit As SAPbouiCOM.EditText
        For nRow = 1 To oMatrix.RowCount
            oEdit = colMark.Cells.Item(nRow).Specific
            oEdit.Value = CStr(nRow)
        Next

    End Sub


    ' This procedure adds all the contact persons of specified business partner code to a combo box
    Private Sub AddContactPersonsCombo(ByVal CardCode As String)
        Dim oCombo As SAPbouiCOM.ComboBox
        oCombo = oShipmentForm.Items.Item(ctlCnctCode).Specific
        Dim nCount As Integer
        nCount = oCombo.ValidValues.Count
        Dim nIndex As Integer
        For nIndex = 1 To nCount
            oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
        Next

        Dim Rs As SAPbobsCOM.Recordset
        Rs = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Rs.DoQuery("SELECT CntctCode, Name FROM OCPR WHERE CardCode = '" & CardCode & "'")
        Rs.MoveFirst()
        While Rs.EoF = False
            oCombo.ValidValues.Add(Rs.Fields.Item("CntctCode").Value, Rs.Fields.Item("Name").Value)
            Rs.MoveNext()
        End While

    End Sub

    ' This procedure adds a new menu item under the "Sales" menu
    Private Sub AddMenuItems()
        Dim oMenus As SAPbouiCOM.Menus          ' The menus collection
        Dim oMenuItem As SAPbouiCOM.MenuItem    ' The new menu item

        ' Get the menus collection from the application
        oMenus = SBOApplication.Menus

        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams

        oCreationPackage = SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

        oMenuItem = SBOApplication.Menus.Item(mnuPurchasing) ' Purhasing menu ID
        oMenus = oMenuItem.SubMenus

        ' New menu parameters
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        oCreationPackage.UniqueID = mnuShipment
        oCreationPackage.String = "Shipment Document"
        oCreationPackage.Enabled = True
        oCreationPackage.Position = 15

        Try ' If the manu already exists this code will fail
            oMenus.AddEx(oCreationPackage)
        Catch er As Exception ' Menu already exists
            'SBOApplication.MessageBox("Menu Already Exists")
        End Try

    End Sub

    Private Sub AddTerritoryField(ByRef oForm As SAPbouiCOM.Form)

        Dim oEditText As SAPbouiCOM.EditText
        Dim oStaticText As SAPbouiCOM.StaticText
        Dim oLinkedButton As SAPbouiCOM.LinkedButton
        Dim oImgButton As SAPbouiCOM.Button
        Dim oItem As SAPbouiCOM.Item
        Dim oRelItem As SAPbouiCOM.Item
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Try

            oCFLs = oForm.ChooseFromLists

            oCFLCreationParams = SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.UniqueID = "cflTerrit1"
            oCFLCreationParams.ObjectType = "200"
            oCFLCreationParams.MultiSelection = False
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCFLCreationParams = Nothing

            oCFLCreationParams = SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.UniqueID = "cflTerrit2"
            oCFLCreationParams.ObjectType = "200"
            oCFLCreationParams.MultiSelection = False
            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding a User Data Source
            oForm.DataSources.UserDataSources.Add("ds1Territ", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("ds2Territ", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oRelItem = oForm.Items.Item("230")

            Dim nHeightGap As Integer
            nHeightGap = (oRelItem.Top - oForm.Items.Item("21").Top) - 1

            oRelItem = oForm.Items.Item("222")

            oItem = oForm.Items.Add("fldTerrit", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = oRelItem.Left
            oItem.Top = oRelItem.Top + nHeightGap
            oItem.Width = oRelItem.Width
            oItem.Height = oRelItem.Height
            oItem.LinkTo = "lblTerrit"

            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "", "ds1Territ")
            oEditText.ChooseFromListUID = "cflTerrit1"
            oEditText.ChooseFromListAlias = "descript"
            oEditText.TabOrder = 1660

            oItem = oForm.Items.Add("bndTerrit", SAPbouiCOM.BoFormItemTypes.it_EDIT)
            oItem.Left = oRelItem.Left - 10
            oItem.Top = oRelItem.Top + nHeightGap + nHeightGap
            oItem.Width = -1 ' Hide this box
            oItem.Height = oRelItem.Height
            oItem.Enabled = False

            oEditText = oItem.Specific
            oEditText.TabOrder = 1660
            If oForm.TypeEx = frmARCreditMemo Then
                oEditText.DataBind.SetBound(True, "ORIN", "U_Territ")
            Else
                oEditText.DataBind.SetBound(True, "OINV", "U_Territ")
            End If

            oRelItem = oForm.Items.Item("230")

            oItem = oForm.Items.Add("lblTerrit", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = oRelItem.Left
            oItem.Top = oRelItem.Top + nHeightGap
            oItem.Width = oRelItem.Width
            oItem.Height = oRelItem.Height
            oItem.LinkTo = "fldTerrit"

            oStaticText = oItem.Specific
            oStaticText.Caption = "Territory"

            oRelItem = oForm.Items.Item("229")

            oItem = oForm.Items.Add("lnkTerrit", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
            oItem.Left = oRelItem.Left
            oItem.Top = oRelItem.Top + nHeightGap
            oItem.Width = oRelItem.Width
            oItem.Height = oRelItem.Height
            oItem.LinkTo = "bndTerrit"

            oLinkedButton = oItem.Specific
            oLinkedButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Territory
            oLinkedButton.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_Territory

            oRelItem = oForm.Items.Item("225")
            Dim oRelLinkImage As SAPbouiCOM.Button
            oRelLinkImage = oRelItem.Specific

            oItem = oForm.Items.Add("btnTerrit", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = oRelItem.Left
            oItem.Top = oRelItem.Top + nHeightGap
            oItem.Width = oRelItem.Width
            oItem.Height = oRelItem.Height
            oItem.LinkTo = "fldTerrit"

            oImgButton = oItem.Specific
            oImgButton.Type = SAPbouiCOM.BoButtonTypes.bt_Image
            oImgButton.Image = "CHOOSE_ICON"
            oImgButton.ChooseFromListUID = "cflTerrit2"

        Catch ex As Exception

            SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Private Sub AddDriverAndTruckField(ByRef oForm As SAPbouiCOM.Form)

        Dim oComboBox As SAPbouiCOM.ComboBox
        Dim oStaticText As SAPbouiCOM.StaticText
        Dim oItem As SAPbouiCOM.Item
        Dim oRelItem As SAPbouiCOM.Item
        Try

            oRelItem = oForm.Items.Item("230")

            Dim nHeightGap As Integer
            nHeightGap = (oRelItem.Top - oForm.Items.Item("21").Top) - 1

            oRelItem = oForm.Items.Item("222")

            oItem = oForm.Items.Add("fldDriver", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)

            oItem.Left = oRelItem.Left
            oItem.Top = oRelItem.Top + nHeightGap
            oItem.Width = oRelItem.Width
            oItem.Height = oRelItem.Height
            oItem.LinkTo = "lblDriver"
            oItem.DisplayDesc = True

            oComboBox = oItem.Specific
            oComboBox.DataBind.SetBound(True, "ODLN", "U_Driver")
            oComboBox.TabOrder = 1660

            Dim Rs As SAPbobsCOM.Recordset
            Rs = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oComboBox.ValidValues.Add("", "")
            Rs.DoQuery("SELECT Code, Name FROM [@MAM_DRV]")
            Rs.MoveFirst()
            While Rs.EoF = False
                oComboBox.ValidValues.Add(Rs.Fields.Item("Code").Value, Rs.Fields.Item("Name").Value)
                Rs.MoveNext()
            End While

            oRelItem = oForm.Items.Item("230")

            oItem = oForm.Items.Add("lDriver", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = oRelItem.Left
            oItem.Top = oRelItem.Top + nHeightGap
            oItem.Width = oRelItem.Width
            oItem.Height = oRelItem.Height
            oItem.LinkTo = "fldDriver"

            oStaticText = oItem.Specific
            oStaticText.Caption = "Driver & Truck"

        Catch ex As Exception

            SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub

    Private Sub SaveAsXML(ByRef Form As SAPbouiCOM.Form)

        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.StartupPath).ToString

        Dim sFileName As String
        sFileName = sPath & "\Form" & Form.TypeEx & ".xml"

        If IO.File.Exists(sFileName) Then
            Exit Sub
        End If

        Dim oXmlDoc As Xml.XmlDocument
        oXmlDoc = New Xml.XmlDocument

        '// get the form as an XML string
        Dim sXmlString As String
        sXmlString = Form.GetAsXML

        '// load the form's XML string to the
        '// XML document object
        oXmlDoc.LoadXml(sXmlString)

        '// save the XML Document
        oXmlDoc.Save(sFileName)

    End Sub

    Private Sub LoadFromXML(ByRef FileName As String)

        Dim oXmlDoc As Xml.XmlDocument

        oXmlDoc = New Xml.XmlDocument

        Dim thisExe As System.Reflection.Assembly
        thisExe = System.Reflection.Assembly.GetExecutingAssembly()

        Dim file As System.IO.Stream

        file = thisExe.GetManifestResourceStream("MamcoAddOns." & FileName)

        '// load the content of the XML File
        'Dim sPath As String

        'sPath = IO.Directory.GetParent(Application.StartupPath).ToString

        oXmlDoc.Load(file) ' sPath & "\" & FileName

        '// load the form to the SBO application in one batch
        SBOApplication.LoadBatchActions(oXmlDoc.InnerXml)

    End Sub

    Private Function GetKeyFromXML(ByVal sObjectKey As String, ByVal sPath As String) As Long

        Dim oXmlDoc As Xml.XmlDocument
        Dim oXmlNode As Xml.XmlNode

        oXmlDoc = New Xml.XmlDocument

        Try
            oXmlDoc.LoadXml(sObjectKey)
            oXmlNode = oXmlDoc.SelectSingleNode(sPath)
            If Not oXmlNode Is Nothing Then
                GetKeyFromXML = Long.Parse(oXmlNode.InnerText)
            End If
        Catch ex As Exception
            GetKeyFromXML = 0
        End Try

    End Function

    Private Sub RefreshForm(ByRef formId As String)
        Try
            Dim f As SAPbouiCOM.Form

            f = SBOApplication.Forms.Item(formId)
            f.Refresh()
        Catch ex As Exception
        End Try
    End Sub


    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.ItemEvent


        Try

        
            ' ============================================================
            '
            ' Purchase Orders / Good Receipts PO Item Events
            '
            ' ============================================================
            If ((pVal.FormTypeEx = frmPurchaseOrders _
                        Or pVal.FormTypeEx = frmGoodsReceiptPO _
                        Or pVal.FormTypeEx = frmInventoryTransfer) _
                    And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD) _
                    And (pVal.Before_Action = True) Then

                ' get the event's form
                Dim oForm As SAPbouiCOM.Form
                oForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                'SaveAsXML(oForm)

                If ((pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD) And _
                    (pVal.Before_Action = True)) Then
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim oOkItem As SAPbouiCOM.Item
                        Dim oCancelItem As SAPbouiCOM.Item
                        Dim oDraftItem As SAPbouiCOM.Item
                        Dim oStatusItem As SAPbouiCOM.Item
                        Dim isDraft As Boolean
                        isDraft = False
                        If Not pVal.FormTypeEx = frmInventoryTransfer Then
                            oStatusItem = oForm.Items.Item("81")
                            If Not oStatusItem Is Nothing Then
                                Dim oStatusCombo As SAPbouiCOM.ComboBox
                                oStatusCombo = oStatusItem.Specific
                                If oStatusCombo.Selected.Value = "6" Then
                                    isDraft = True
                                Else
                                    isDraft = False
                                End If
                            End If
                        Else
                            If InStr(oForm.Title, " - Draft") > 0 Then
                                isDraft = True
                            End If
                        End If
                        oOkItem = oForm.Items.Item("1")
                        oCancelItem = oForm.Items.Item("2")
                        oDraftItem = oForm.Items.Add(btnDraft, SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                        oDraftItem.Width = oOkItem.Width
                        oDraftItem.Height = oOkItem.Height
                        oDraftItem.Top = oOkItem.Top
                        oDraftItem.Left = oCancelItem.Left + oCancelItem.Width + (oCancelItem.Left - (oOkItem.Left + oOkItem.Width))
                        oDraftBtn = oDraftItem.Specific
                        If isDraft Then
                            oDraftBtn.Caption = "Update Draft"
                        Else
                            oDraftBtn.Caption = "Save Draft"
                        End If

                        Dim nTempLeft As Integer
                        ' swap ok with save draft
                        nTempLeft = oOkItem.Left
                        oOkItem.Left = oDraftItem.Left
                        oDraftItem.Left = nTempLeft
                        ' swap ok with cancel button
                        nTempLeft = oCancelItem.Left
                        oCancelItem.Left = oOkItem.Left
                        oOkItem.Left = nTempLeft
                        ' make draft default button
                        oForm.DefButton = oDraftItem.UniqueID

                    End If
                End If

                If ((pVal.ItemUID = btnDraft) And _
                    (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)) Then
                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    ' Execute Save as Draft command if is enabled
                    oMenuItem = SBOApplication.Menus.Item("5907")
                    If oMenuItem.Enabled = True Then
                        Dim nDocKey As Integer
                        nDocKey = GetKeyFromXML(oForm.BusinessObject.Key, "DocumentParams/DocEntry")
                        SBOApplication.ActivateMenuItem("5907")
                    Else
                        SBOApplication.MessageBox("Currently, this document cannot be saved as draft!")
                    End If
                End If

                If ((pVal.ItemUID = "1") _
                        And (pVal.FormMode = 3) _
                        And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) _
                        And (pVal.BeforeAction = True)) Then
                    AddStarted = True
                    If (pVal.FormMode = 3) Then
                        nDraftKey = 0
                        Dim isDraft As Boolean
                        isDraft = False
                        Dim sConfirmMessage As String
                        sConfirmMessage = "Are you sure you want to save it as normal and not draft?"
                        If Not pVal.FormTypeEx = frmInventoryTransfer Then
                            Dim oStatusItem As SAPbouiCOM.Item
                            oStatusItem = oForm.Items.Item("81")
                            If Not oStatusItem Is Nothing Then
                                Dim oStatusCombo As SAPbouiCOM.ComboBox
                                oStatusCombo = oStatusItem.Specific
                                If oStatusCombo.Selected.Value = "6" Then ' Is it a draft document?
                                    sConfirmMessage = "Are you sure you want change the draft document to normal?"
                                    isDraft = True
                                End If
                            End If
                        Else
                            If InStr(oForm.Title, " - Draft") > 0 Then
                                sConfirmMessage = "Are you sure you want change the draft document to normal?"
                                isDraft = True
                            End If
                        End If

                        If SBOApplication.MessageBox(sConfirmMessage, 2, "Yes", "No") = 2 Then
                            BubbleEvent = False
                            AddStarted = False
                        Else
                            If isDraft Then
                                ' keep draft key to remove it after saving document
                                If pVal.FormTypeEx = frmInventoryTransfer Then
                                    Dim oDocNumItem As SAPbouiCOM.EditText
                                    oDocNumItem = oForm.Items.Item("11").Specific
                                    Dim Rs As SAPbobsCOM.Recordset
                                    Rs = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Rs.DoQuery("SELECT DocEntry FROM ODRF WHERE DocNum = " & oDocNumItem.Value & " AND DocType = 'I'")
                                    If Not Rs.EoF Then
                                        nDraftKey = Rs.Fields.Item("DocEntry").Value
                                    End If
                                    Rs = Nothing
                                    System.GC.Collect()
                                Else
                                    nDraftKey = GetKeyFromXML(oForm.BusinessObject.Key, "DocumentParams/DocEntry")
                                End If
                            End If
                        End If
                    End If
                End If

                If ((pVal.ItemUID = btnDraft) _
                        And (pVal.FormMode = 3) _
                        And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)) Then
                    AddStarted = False
                    '    If (pVal.FormMode = 3) And _
                    '        SBOApplication.MessageBox("Are you sure you want to save it as normal and not draft?", 2, "Yes", "No") = 2 Then
                    '        BubbleEvent = False
                    '    End If
                End If

            End If

            ' ============================================================
            '
            ' Pick List Form Item Events
            '
            ' ============================================================
            If (pVal.FormTypeEx = frmPickList) And (pVal.FormMode = 2) Then
                ' Is Picked Qty changed?
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE _
                        And (pVal.ItemUID = "11") _
                        And (pVal.ColUID = "19") _
                        And (pVal.Before_Action = False) _
                        And (pVal.Action_Success = True) Then
                    If pVal.Before_Action Then
                        PickQtyChanged = False
                    Else
                        PickQtyChanged = pVal.ItemChanged
                    End If
                End If

                ' Is Picked Qty changed?
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS _
                        And (pVal.ItemUID = "11") _
                        And (pVal.ColUID = "19") _
                        And (pVal.Before_Action = False) _
                        And (pVal.Action_Success = True) Then

                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                    Dim oMatrix As SAPbouiCOM.Matrix
                    oMatrix = oForm.Items.Item("11").Specific

                    Dim oPicked As SAPbouiCOM.EditText
                    oPicked = oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific
                    Dim oBinNo1 As SAPbouiCOM.EditText
                    oBinNo1 = oMatrix.Columns.Item("U_BinNo1").Cells.Item(pVal.Row).Specific
                    If Len(oPicked.Value) > 0 And Len(oBinNo1.Value) = 0 Then
                        Dim oItemCode As SAPbouiCOM.EditText
                        oItemCode = oMatrix.Columns.Item("12").Cells.Item(pVal.Row).Specific
                        Dim nPickedQty As Double
                        nPickedQty = Convert.ToDouble(oPicked.Value, NumberProvider)
                        'nPickedQty = SBOFunctions.CleanNumberAsDouble(oPicked.Value)
                        Dim nRemainQty As Double
                        nRemainQty = nPickedQty
                        Dim nBinIndex As Integer
                        nBinIndex = 1
                        Dim Rs As SAPbobsCOM.Recordset
                        Rs = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Rs.DoQuery("SELECT BatchNum, Quantity FROM OIBT WHERE ItemCode = '" & oItemCode.Value & "' AND WhsCode = '01' AND Quantity > 0")
                        Rs.MoveFirst()
                        While Rs.EoF = False
                            Dim sBinNo As String
                            Dim nBinQty As Double
                            sBinNo = Rs.Fields.Item("BatchNum").Value
                            nBinQty = Rs.Fields.Item("Quantity").Value
                            If nBinQty > 0 Then
                                Dim oBinNo As SAPbouiCOM.EditText
                                oBinNo = oMatrix.Columns.Item("U_BinNo" & nBinIndex).Cells.Item(pVal.Row).Specific
                                Try
                                    oBinNo.Value = sBinNo
                                Catch ex As Exception
                                    SBOApplication.StatusBar.SetText("Error setting Bin No (" & nBinIndex & "): " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                                Dim oBinQty As SAPbouiCOM.EditText
                                oBinQty = oMatrix.Columns.Item("U_BinQty" & nBinIndex).Cells.Item(pVal.Row).Specific
                                Try
                                    If nBinQty > nRemainQty Then
                                        oBinQty.Value = CStr(nRemainQty)
                                        Exit While
                                    End If
                                    oBinQty.Value = CStr(nBinQty)
                                Catch ex As Exception
                                    SBOApplication.StatusBar.SetText("Error setting Bin Qty (" & nBinIndex & "): " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                                nRemainQty = Math.Max(nRemainQty - nBinQty, 0)
                                nBinIndex = nBinIndex + 1
                            End If
                            Rs.MoveNext()
                        End While
                        Rs = Nothing
                        System.GC.Collect()

                        PickQtyChanged = False
                        '// Get the DBdatasource we base the matrix on
                        'Dim oDBDataSource As SAPbouiCOM.DBDataSource
                        'oDBDataSource = oForm.DataSources.DBDataSources.Item("PKL1")
                        'Try
                        '   oDBDataSource.SetValue("U_BinNo1", pVal.Row - 1, "bin1")
                        '   oDBDataSource.SetValue("U_BinQty1", pVal.Row - 1, oPicked.Value)
                        'Catch ex As Exception
                        '    SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'End Try
                    End If
                End If

            End If

            ' ============================================================
            '
            ' Batch Selection Item Events
            '
            ' ============================================================
            If (pVal.FormTypeEx = frmBatchSelection) Then

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD _
                        And (pVal.Before_Action = False)) _
                        Or (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK _
                            And (pVal.ItemUID = "3") _
                            And (pVal.Before_Action = False)) Then

                    isBatchUpdatedAutomatically = False

                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                    'SaveAsXML(oForm)

                    Dim oItemMatrix As SAPbouiCOM.Matrix
                    oItemMatrix = oForm.Items.Item("3").Specific

                    Dim isMainWarehouse As Boolean
                    isMainWarehouse = False
                    Dim oWarehouse As SAPbouiCOM.EditText
                    oWarehouse = oItemMatrix.Columns.Item("3").Cells.Item(1).Specific()
                    If oWarehouse.Value = "01" Then
                        isMainWarehouse = True
                    End If

                    If Not isMainWarehouse And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then

                        Dim oAutoSelect As SAPbouiCOM.Item
                        oAutoSelect = oForm.Items.Item("16")

                        Dim oUpdateButton As SAPbouiCOM.Item
                        oUpdateButton = oForm.Items.Item("1")

                        Dim isAutoSelected As Boolean
                        isAutoSelected = False

                        Try
                            Dim nRow As Integer
                            For nRow = 1 To oItemMatrix.RowCount
                                If oItemMatrix.IsRowSelected(nRow) Then
                                    If oAutoSelect.Enabled Then
                                        oAutoSelect.Click()
                                        oUpdateButton.Click()

                                        'Dim oWarehouse As SAPbouiCOM.EditText
                                        'oWarehouse = oItemMatrix.Columns.Item("3").Cells.Item(nRow).Specific()
                                        'If oWarehouse.Value = "01" Then
                                        '    isMainWarehouse = True
                                        'End If

                                        Dim nNextRow As Integer
                                        For nNextRow = nRow + 1 To oItemMatrix.RowCount
                                            Try
                                                oItemMatrix.Columns.Item("55").Cells.Item(nNextRow).Click()
                                                'oItemMatrix.SelectRow(nNextRow, True, False)
                                                Exit For
                                            Catch ex As Exception
                                                ' Ignore this, try next one
                                            End Try
                                        Next
                                    End If
                                End If
                            Next
                            isAutoSelected = True
                        Catch ex As Exception
                            SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try

                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            If Not isMainWarehouse Then
                                Try
                                    isBatchUpdatedAutomatically = True
                                    oForm.Close()
                                    SBOApplication.MessageBox("Batches updated automatically! Please, press [Add] button again.")
                                    'SBOApplication.StatusBar.SetText("Batches updated automatically. Please, press Add button again!", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Catch ex As Exception
                                    SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                            Else
                                SBOApplication.StatusBar.SetText("Batches updated automatically. Confirm changes and press OK button!", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        Else
                            SBOApplication.StatusBar.SetText("Could not update all batches automatically. You have to update batches manually!", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If

                    ElseIf isMainWarehouse Then

                        Dim nSelectedRow As Integer
                        If pVal.ItemUID = "3" Then
                            nSelectedRow = pVal.Row
                        Else
                            Dim nRow As Integer
                            For nRow = 1 To oItemMatrix.RowCount
                                If oItemMatrix.IsRowSelected(nRow) Then
                                    nSelectedRow = nRow
                                    Exit For
                                End If
                            Next
                        End If

                        ' Make details editable
                        'Try
                        '   oBatchMatrix.Columns.Item("11").Editable = True
                        'Catch ex As Exception
                        '   SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'End Try

                        Dim oBatchMatrix As SAPbouiCOM.Matrix
                        oBatchMatrix = oForm.Items.Item("4").Specific

                        Dim oBaseEntry As SAPbouiCOM.EditText
                        oBaseEntry = oItemMatrix.Columns.Item("58").Cells.Item(nSelectedRow).Specific

                        Dim oBaseLine As SAPbouiCOM.EditText
                        oBaseLine = oItemMatrix.Columns.Item("62").Cells.Item(nSelectedRow).Specific

                        Dim oQtyNeeded As SAPbouiCOM.EditText
                        oQtyNeeded = oItemMatrix.Columns.Item("55").Cells.Item(nSelectedRow).Specific

                        Dim nQtyNeeded As Double
                        nQtyNeeded = Convert.ToDouble(oQtyNeeded.Value, NumberProvider)
                        'nQtyNeeded = SBOFunctions.CleanNumberAsDouble(oQtyNeeded.Value)

                        Dim oSelectBatch As SAPbouiCOM.Item
                        oSelectBatch = oForm.Items.Item("48")

                        If nQtyNeeded > 0 Then

                            Dim Rs As SAPbobsCOM.Recordset
                            Rs = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Rs.DoQuery("SELECT U_BinNo1, U_BinQty1, U_BinNo2, U_BinQty2, U_BinNo3, U_BinQty3 FROM PKL1 WHERE OrderEntry = " & oBaseEntry.Value & " AND OrderLine = " & oBaseLine.Value)
                            Rs.MoveFirst()
                            If Not Rs.EoF Then
                                Dim nBinIndex As Integer
                                For nBinIndex = 1 To 3
                                    Dim sBinNo As String
                                    Dim nBinQty As Double
                                    sBinNo = Rs.Fields.Item("U_BinNo" & nBinIndex).Value
                                    nBinQty = CDbl(Rs.Fields.Item("U_BinQty" & nBinIndex).Value)
                                    If sBinNo <> "" And nBinQty > 0 Then
                                        Dim nRow As Integer
                                        For nRow = 1 To oBatchMatrix.RowCount
                                            Dim oBinNo As SAPbouiCOM.EditText
                                            oBinNo = oBatchMatrix.Columns.Item("0").Cells.Item(nRow).Specific
                                            Dim oQtySelected As SAPbouiCOM.EditText
                                            oQtySelected = oBatchMatrix.Columns.Item("4").Cells.Item(nRow).Specific
                                            If oBinNo.Value = sBinNo Then
                                                Try
                                                    oQtySelected.Value = CStr(nBinQty)
                                                    oSelectBatch.Click()
                                                    'oBatchMatrix.SelectRow(nRow, True, False)
                                                Catch ex As Exception
                                                    SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                End Try
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                            Rs = Nothing
                            System.GC.Collect()

                        End If
                    End If

                End If

            End If

            ' ============================================================
            '
            ' Delivery Item Events
            '
            ' ============================================================
            If (pVal.FormTypeEx = frmDelivery) Then

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD _
                        And (pVal.Before_Action = False)) Then

                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                    'SaveAsXML(oForm)
                    AddDriverAndTruckField(oForm)

                End If

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED _
                        And pVal.ItemUID = "1" _
                        And (pVal.FormMode = 2 Or pVal.FormMode = 3) _
                        And (pVal.Before_Action = True)) Then

                    Dim oUDFForm As SAPbouiCOM.Form
                    oUDFForm = SBOApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    Dim oComboBox As SAPbouiCOM.ComboBox
                    oComboBox = oUDFForm.Items.Item("fldDriver").Specific
                    If oComboBox.Selected.Value = "" Then
                        SBOApplication.StatusBar.SetText("Field 'Driver & Truck' is required!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False
                    End If

                End If

            End If

            ' ============================================================
            '
            ' Business Partner
            '
            ' ============================================================
            If (pVal.FormTypeEx = frmBusinessPartner) Then

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD _
                        And (pVal.Before_Action = False)) Then

                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                    'SaveAsXML(oForm)

                End If


            End If


            ' ============================================================
            '
            ' Landed Costs
            '
            ' ============================================================
            If (pVal.FormTypeEx = frmLandedCosts) Then

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD _
                        And (pVal.Before_Action = False)) Then

                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                    'SaveAsXML(oForm)

                    'Try

                    '    oForm.DataSources.UserDataSources.Add("dsCardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 15)
                    '    oForm.DataSources.UserDataSources.Add("dsCardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)

                    '    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
                    '    Dim oCFL As SAPbouiCOM.ChooseFromList
                    '    Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams

                    '    oCFLs = oForm.ChooseFromLists

                    '    oCFLCreationParams = SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
                    '    oCFLCreationParams.UniqueID = "cflSupplier"
                    '    oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner
                    '    oCFLCreationParams.MultiSelection = False

                    '    oCFL = oCFLs.Add(oCFLCreationParams)

                    '    Dim oCons As SAPbouiCOM.Conditions
                    '    oCons = oCFL.GetConditions()

                    '    Dim oCon As SAPbouiCOM.Condition
                    '    oCon = oCons.Add()
                    '    oCon.Alias = "CardType"
                    '    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    '    oCon.CondVal = "S"
                    '    oCFL.SetConditions(oCons)

                    '    oCFLCreationParams = Nothing

                    '    Dim oMatrix As SAPbouiCOM.Matrix
                    '    oMatrix = oForm.Items.Item("3").Specific
                    '    oMatrix.Clear()

                    '    Dim oColumn As SAPbouiCOM.Column
                    '    oColumn = oMatrix.Columns.Add("Supplier", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                    '    oColumn.Width = 10
                    '    oColumn.Editable = True
                    '    oColumn.TitleObject.Caption = "Supplier"
                    '    'oColumn.DataBind.SetBound(True, "@MAM_CRGS", "U_SCode")
                    '    oColumn.DataBind.SetBound(True, "", "dsCardName")

                    '    oColumn.ChooseFromListUID = "cflSupplier"
                    '    oColumn.ChooseFromListAlias = "CardCode"

                    '    Dim oLink As SAPbouiCOM.LinkedButton
                    '    oLink = oColumn.ExtendedObject
                    '    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner
                    '    oLink.LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner

                    '    oMatrix.AddRow(1, -1)

                    'Catch ex As Exception
                    '    SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    'End Try

                End If

            End If

            ' ============================================================
            '
            ' Delivery Item Events
            '
            ' ============================================================
            If (pVal.FormTypeEx = frmItemMaster) Then

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD _
                        And (pVal.Before_Action = False)) Then

                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                    'SaveAsXML(oForm)

                End If

            End If


            ' ============================================================
            '
            ' Pick Manager Item Events
            '
            ' ============================================================
            If (pVal.FormTypeEx = frmPickManager) Then

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED _
                        And (pVal.ItemUID = "11") _
                        And (pVal.Before_Action = False) _
                        And (pVal.Action_Success = True)) Then

                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                    'SaveAsXML(oForm)

                End If

            End If

            ' ============================================================
            '
            ' Pick Details Events
            '
            ' ============================================================
            If (pVal.FormTypeEx = frmPickDetails) Then

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD _
                        And (pVal.Before_Action = False)) Then

                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                    'SaveAsXML(oForm)

                End If

            End If


            ' ============================================================
            '
            ' Pick Details Events
            '
            ' ============================================================
            If (pVal.FormTypeEx = frmIncomingPayments _
                    Or pVal.FormTypeEx = frmOutgoingPayments) Then

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD _
                        And (pVal.Before_Action = False)) Then

                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBOApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    'SaveAsXML(oForm)

                End If

            End If


            ' ============================================================
            '
            ' A/R Invoice + Payment
            '
            ' ============================================================
            If (pVal.FormTypeEx = frmARInvoicePayment) Then

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD _
                        And (pVal.Before_Action = False)) Then
                    isARInvoicePaymentLoading = True
                End If

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN _
                        And (pVal.ItemUID = "38") _
                        And (pVal.ColUID = "11") _
                        And (pVal.CharPressed = 40) _
                        And (pVal.Before_Action = False)) Then

                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                    Dim oMatrix As SAPbouiCOM.Matrix
                    oMatrix = oForm.Items.Item("38").Specific

                    If pVal.Row = oMatrix.RowCount - 1 Then

                        Try
                            oMatrix.Columns.Item("4").Cells.Item(pVal.Row + 1).Click()
                        Catch ex As Exception
                            SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try

                    End If

                End If

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE _
                        And (pVal.Before_Action = False)) Then

                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                    'SaveAsXML(oForm)

                    SBOApplication.StatusBar.SetText("Form is Activated!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_None)

                    Dim oMatrix As SAPbouiCOM.Matrix
                    oMatrix = oForm.Items.Item("38").Specific

                    If isARInvoicePaymentLoading Then
                        isARInvoicePaymentLoading = False
                        Try
                            oMatrix.Columns.Item("4").Cells.Item(1).Click()
                        Catch ex As Exception
                            SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

                    'If isBatchUpdatedAutomatically Then
                    '    isBatchUpdatedAutomatically = False
                    '    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    '        Try
                    '            oForm.Items.Item("1").Click()
                    '        Catch ex As Exception
                    '            SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        End Try
                    '    End If
                    'End If

                End If

            End If

            ' ============================================================
            '
            ' A/R Invoice ...
            '
            ' ============================================================
            If (pVal.FormTypeEx = frmARInvoice) _
                    Or (pVal.FormTypeEx = frmARInvoicePayment) _
                    Or (pVal.FormTypeEx = frmARCreditMemo) _
                    Or (pVal.FormTypeEx = frmARReverseInvoice) Then

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD _
                        And (pVal.Before_Action = False)) Then

                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                    AddTerritoryField(oForm)

                    'SaveAsXML(oForm)

                    If pVal.FormType = frmARReverseInvoice Then

                        Dim oRS As SAPbobsCOM.Recordset
                        oRS = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRS.DoQuery("SELECT ICTCard FROM ousr u, oudg d WHERE u.DfltsGroup = d.Code AND u.user_code = '" & SBOCompany.UserName & "'")
                        oRS.MoveFirst()
                        If Not oRS.EoF Then
                            Dim oEditText As SAPbouiCOM.EditText

                            Try
                                oEditText = oForm.Items.Item("4").Specific
                                oEditText.Value = oRS.Fields.Item("ICTCard").Value
                                oForm.Items.Item("4").Enabled = False
                                oForm.Items.Item("67").Enabled = False

                            Catch ex As Exception
                                SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                        End If

                    End If


                End If

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS _
                        And (pVal.ItemUID = "4") _
                        And (pVal.Before_Action = False) _
                        And (pVal.Action_Success = True) Then

                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)
                    Dim sTerritId As String
                    sTerritId = ""
                    Try
                        Dim sTerritDesc As String
                        sTerritId = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("Territory", 0)
                        If sTerritId <> "" Then
                            Dim oRs As SAPbobsCOM.Recordset
                            oRs = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs.DoQuery("SELECT descript FROM OTER WHERE territryID = " & sTerritId)
                            oRs.MoveFirst()
                            If Not oRs.EoF Then
                                sTerritDesc = oRs.Fields.Item("descript").Value
                            Else
                                sTerritDesc = ""
                            End If
                        Else
                            sTerritDesc = ""
                        End If
                        oForm.DataSources.UserDataSources.Item("ds1Territ").ValueEx = sTerritDesc
                        oForm.DataSources.UserDataSources.Item("ds2Territ").ValueEx = sTerritId
                    Catch ex As Exception
                        SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try

                    Try
                        Dim oEdit As SAPbouiCOM.EditText
                        oEdit = oForm.Items.Item("bndTerrit").Specific
                        oEdit.Value = sTerritId
                    Catch ex As Exception
                        SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                    Try
                        Dim oUDFForm As SAPbouiCOM.Form
                        oUDFForm = SBOApplication.Forms.GetForm("-" & pVal.FormType, pVal.FormTypeCount)
                        Dim oEdit As SAPbouiCOM.EditText
                        oEdit = oUDFForm.Items.Item("U_Territ").Specific
                        oEdit.Value = sTerritId
                    Catch ex As Exception
                        ' ignore
                    End Try
                End If

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED _
                        And pVal.ItemUID = "1" _
                        And (pVal.FormMode = 2 Or pVal.FormMode = 3) _
                        And (pVal.Before_Action = True)) Then

                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                    Dim oEdit As SAPbouiCOM.EditText
                    oEdit = oForm.Items.Item("bndTerrit").Specific
                    If oEdit.Value = "" Then
                        SBOApplication.StatusBar.SetText("Field Territory is required!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False
                    End If

                End If

            End If

            ' ============================================================
            '
            ' CFL Territories
            '
            ' ============================================================
            If (pVal.FormTypeEx = frmCFLTerritories) Then

                If (Not oCFLForm Is Nothing) _
                        And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD _
                        And (pVal.Before_Action = False)) Then

                    Dim sTerritId As String
                    sTerritId = oCFLForm.DataSources.UserDataSources.Item("ds2Territ").ValueEx()
                    Try
                        Dim oTerrit As SAPbouiCOM.EditText
                        oTerrit = oCFLForm.Items.Item("bndTerrit").Specific
                        oTerrit.Value = sTerritId
                    Catch ex As Exception
                        SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try

                    Try
                        Dim oUDFForm As SAPbouiCOM.Form
                        oUDFForm = SBOApplication.Forms.GetForm("-" & oCFLForm.TypeEx, oCFLForm.TypeCount)
                        Dim oUDFTerrit As SAPbouiCOM.EditText
                        oUDFTerrit = oUDFForm.Items.Item("U_Territ").Specific
                        oUDFTerrit.Value = sTerritId
                    Catch ex As Exception

                    End Try


                    'Dim oForm As SAPbouiCOM.Form
                    'oForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                    'Try
                    '    Dim sTerritId As String
                    '    sTerritId = oForm.DataSources.UserDataSources.Item("ds2Territ").ValueEx
                    '    Dim sDBTerritId As String
                    '    If sTerritId <> sDBTerritId Then
                    '        Dim oEdit As SAPbouiCOM.EditText
                    '        oEdit = oForm.Items.Item("bndTerrit").Specific
                    '        oEdit.Value = sTerritId
                    '    End If
                    'Catch ex As Exception
                    '    SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    'End Try

                End If

            End If


            ' ============================================================
            '
            ' CFL Suppliers
            '
            ' ============================================================
            If (pVal.FormTypeEx = frmCFLSuppliers) Then

                If (Not oCFLForm Is Nothing) _
                        And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD _
                        And (pVal.Before_Action = False)) Then

                    Dim sCardCode As String
                    sCardCode = oCFLForm.DataSources.UserDataSources.Item("dsCardCode").ValueEx
                    Dim sCardName As String
                    sCardName = oCFLForm.DataSources.UserDataSources.Item("dsCardName").ValueEx

                    Try
                    Catch ex As Exception
                        SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try

                End If

            End If

            If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST) Then

                Dim oCFLEvent As SAPbouiCOM.IChooseFromListEvent
                oCFLEvent = pVal

                Dim sCFLId As String
                sCFLId = oCFLEvent.ChooseFromListUID

                oCFLForm = Nothing

                If (sCFLId = "cflTerrit1" Or sCFLId = "cflTerrit2") _
                       And oCFLEvent.BeforeAction = False Then

                    'Dim oTerrit As SAPbouiCOM.EditText
                    'oTerrit = oForm.Items.Item("bndTerrit").Specific
                    'Dim oUDFTerrit As SAPbouiCOM.EditText
                    'oUDFTerrit = Nothing
                    'Try
                    '    Dim oUDFForm As SAPbouiCOM.Form
                    '    oUDFForm = SBOApplication.Forms.GetForm("-" & pVal.FormType, pVal.FormTypeCount)
                    '    oUDFTerrit = oUDFForm.Items.Item("U_Territ").Specific
                    'Catch ex As Exception
                    '    oUDFTerrit = Nothing
                    'End Try

                    Try
                        Dim oDataTable As SAPbouiCOM.DataTable
                        oDataTable = oCFLEvent.SelectedObjects
                        If Not oDataTable Is Nothing Then

                            oCFLForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                            Dim sTerritId As String
                            sTerritId = oDataTable.GetValue("territryID", 0)
                            Dim sTerritDesc As String
                            sTerritDesc = oDataTable.GetValue("descript", 0)
                            oCFLForm.DataSources.UserDataSources.Item("ds1Territ").ValueEx = sTerritDesc
                            oCFLForm.DataSources.UserDataSources.Item("ds2Territ").ValueEx = sTerritId
                            'If sCFLId = "cflTerrit2" Then
                            'oTerrit.Value = sTerritId
                            'If Not oUDFTerrit Is Nothing Then
                            'oUDFTerrit.Value = sTerritId
                            'End If
                            'End If
                        End If
                    Catch ex As Exception
                        SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try

                End If

                'If (sCFLId = "cflSupplier") _
                '       And oCFLEvent.BeforeAction = False Then

                '    Try
                '        Dim oDataTable As SAPbouiCOM.DataTable
                '        oDataTable = oCFLEvent.SelectedObjects
                '        If Not oDataTable Is Nothing Then
                '            oCFLForm = SBOApplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)

                '            Dim sCode As String
                '            sCode = oDataTable.GetValue("CardCode", 0)
                '            Dim sName As String
                '            sName = oDataTable.GetValue("CardName", 0)
                '            oCFLForm.DataSources.UserDataSources.Item("dsCardCode").ValueEx = sCode
                '            oCFLForm.DataSources.UserDataSources.Item("dsCardName").ValueEx = sName
                '        End If
                '    Catch ex As Exception
                '        SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    End Try

                'End If

            End If
        Catch ex As Exception
            SBOApplication.MessageBox(ex.Message)
        End Try
    End Sub
    Private Sub SBO_Application_ToEnableRenameToItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.ItemEvent
        Try
            If FormUID = frmShipment Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                    Dim oCFLEvent As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvent = pVal
                    Dim sCFLId As String
                    sCFLId = oCFLEvent.ChooseFromListUID
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    oCFL = oShipmentForm.ChooseFromLists.Item(sCFLId)

                    If (pVal.Before_Action = False) Then

                        If (sCFLId = "2") Or (sCFLId = "4") Then

                            Dim sVenCode As String
                            Dim sVenName As String
                            Dim sCnctCode As String
                            Dim oDataTable As SAPbouiCOM.DataTable
                            oDataTable = oCFLEvent.SelectedObjects
                            If Not oDataTable Is Nothing Then
                                sVenCode = oDataTable.GetValue("CardCode", 0)
                                sVenName = oDataTable.GetValue("CardName", 0)
                                sCnctCode = oDataTable.GetValue("CntctPrsn", 0)
                                Dim oDBDataSource As SAPbouiCOM.DBDataSource
                                oDBDataSource = oShipmentForm.DataSources.DBDataSources.Item("@MAM_OSHP")
                                oDBDataSource.SetValue("U_VenCode", 0, sVenCode)
                                oDBDataSource.SetValue("U_VenName", 0, sVenName)
                                oDBDataSource.SetValue("U_CnctCode", 0, sCnctCode)
                                AddContactPersonsCombo(sVenCode)
                                Dim oComboBox As SAPbouiCOM.ComboBox
                                oComboBox = oShipmentForm.Items.Item(ctlCnctCode).Specific
                                oComboBox.Select(sCnctCode)
                                SetListCondition(oCFLs.Item("5"), "CardCode", sVenCode)
                                SetListCondition(oCFLs.Item("7"), "CardCode", sVenCode)
                                EnableCopyItemsButton()
                            End If

                        ElseIf (sCFLId = "5") Then

                            Dim sCnctCode As String
                            Dim oDataTable As SAPbouiCOM.DataTable
                            oDataTable = oCFLEvent.SelectedObjects
                            If Not oDataTable Is Nothing Then
                                sCnctCode = oDataTable.GetValue("Name", 0)
                                Dim oComboBox As SAPbouiCOM.ComboBox
                                oComboBox = oShipmentForm.Items.Item(ctlCnctCode).Specific
                                oComboBox.Select(sCnctCode)
                            End If

                        ElseIf (sCFLId = "7") Then

                            Dim sCnctCode As String
                            Dim oDataTable As SAPbouiCOM.DataTable
                            oDataTable = oCFLEvent.SelectedObjects
                            If Not oDataTable Is Nothing Then
                                Dim nOrders As Integer
                                Dim nIndex As Integer
                                nOrders = oDataTable.Rows.Count
                                ReDim OrderCodes(nOrders - 1)
                                For nIndex = 0 To nOrders - 1
                                    OrderCodes(nIndex) = oDataTable.GetValue("DocEntry", nIndex)
                                Next
                                For nIndex = 0 To nOrders - 1

                                Next
                            End If

                        End If
                        'Dim oDataTable As SAPbouiCOM.DataTable
                        'oDataTable = oCFLEvento.SelectedObjects
                        'If Not oDataTable Is Nothing Then
                        '    Dim oCols As SAPbouiCOM.DataColumns
                        '    oCols = oDataTable.Columns
                        '    Dim val As String
                        '    val = oDataTable.GetValue("CardCode", 0)
                        '    Try
                        '        Dim nCol As Integer
                        '        Dim oCol As SAPbouiCOM.DataColumn
                        '        Dim oItem As SAPbouiCOM.Item
                        '        Dim sVendorCode As String
                        '        oCols = oDataTable.Columns
                        '        For nCol = 0 To oCols.Count - 1
                        '            oCol = oCols.Item(nCol)
                        '            If (sCFL_ID = "2" And oCol.Name = "CardCode") Or (sCFL_ID = "4" And oCol.Name = "CardCode") Then
                        '                oItem = oForm.Items.Item(ctlVenCode)
                        '                Dim oEditText As SAPbouiCOM.EditText
                        '                oEditText = oItem.Specific
                        '                Dim val As String
                        '                val = oDataTable.GetValue(oCol.Name, 0)
                        '                sVendorCode = val
                        '                Dim oDBDataSource As SAPbouiCOM.DBDataSource
                        '                oDBDataSource = oForm.DataSources.DBDataSources.Item("@MAM_OSHP")
                        '                oDBDataSource.SetValue("U_VenCode", 0, val)
                        '                If sVendorCode.Length > 0 Then
                        '                    EnableCopyItemsButton()
                        '                End If
                        '                'Try
                        '                'oEditText.Value = val
                        '                'Catch ex As Exception
                        '                'End Try
                        '                oItem = oForm.Items.Item(ctlCnctCode)
                        '                Dim oComboBox As SAPbouiCOM.ComboBox
                        '                oComboBox = oItem.Specific
                        '                AddContactPersonsCombo(oComboBox, sVendorCode)
                        '                SetListCondition(oCFLs.Item("5"), "CardCode", sVendorCode)
                        '            ElseIf (sCFL_ID = "2" And oCol.Name = "CardName") Or (sCFL_ID = "4" And oCol.Name = "CardName") Then
                        '                oItem = oForm.Items.Item(ctlVenName)
                        '                Dim oEditText As SAPbouiCOM.EditText
                        '                oEditText = oItem.Specific
                        '                Dim val As String
                        '                val = oDataTable.GetValue(oCol.Name, 0)
                        '                Dim oDBDataSource As SAPbouiCOM.DBDataSource
                        '                oDBDataSource = oForm.DataSources.DBDataSources.Item("@MAM_OSHP")
                        '                oDBDataSource.SetValue("U_VenName", 0, val)
                        '                'oEditText.Value = val
                        '            ElseIf (sCFL_ID = "2" And oCol.Name = "CntctPrsn") Or (sCFL_ID = "4" And oCol.Name = "CntctPrsn") Or (sCFL_ID = "5" And oCol.Name = "Name") Then
                        '                Dim val As String
                        '                val = oDataTable.GetValue(oCol.Name, 0)
                        '                oItem = oForm.Items.Item(ctlCnctCode)
                        '                Dim oComboBox As SAPbouiCOM.ComboBox
                        '                oComboBox = oItem.Specific
                        '                oComboBox.Select(val)
                        '            End If

                        '        Next

                    Else ' (pVal.Before_Action = True)

                        If (sCFLId = "5") Then

                            Dim sVenCode As String
                            Dim oDBDataSource As SAPbouiCOM.DBDataSource
                            oDBDataSource = oShipmentForm.DataSources.DBDataSources.Item("@MAM_OSHP")
                            sVenCode = oDBDataSource.GetValue("U_VenCode", 0).Trim()
                            If sVenCode.Trim().Length = 0 Then
                                BubbleEvent = False
                            End If

                        End If

                    End If

                    ' Click on Add Row
                    'If (pVal.ItemUID = "AddRow") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) Then
                    '    Dim f As SAPbouiCOM.Form
                    '    Dim oMatrix As SAPbouiCOM.Matrix
                    '    f = SBOApplication.Forms.Item(FormUID)
                    '    oMatrix = f.Items.Item(ctlMatrix).Specific
                    '    f.DataSources.DBDataSources.Item(1).Clear()
                    '    oMatrix.AddRow(1)
                    'End If

                    ' After selecting a BP Code from the combo box
                    'If (pVal.ItemUID = "txtCode") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) Then
                    '    txtBPName.Value = cmbBPCode.Selected.Description
                    'End If

                    ' After selecting an item from the combo box
                    'If (pVal.ItemUID = "mat") And (pVal.ColUID = "ItemCode") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) Then

                    '    Dim oEdit As SAPbouiCOM.EditText
                    '    Dim oCombo As SAPbouiCOM.ComboBox

                    '    oCombo = colItemCode.Cells.Item(pVal.Row).Specific
                    '    oEdit = colItemName.Cells.Item(pVal.Row).Specific
                    '    oEdit.Value = oCombo.Selected.Description
                    'End If

                    ' After changing the item quantity
                    'If (pVal.ItemUID = "mat") And (pVal.ColUID = "ItemQuan") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE) Then
                    '    Dim oEditPrice As SAPbouiCOM.EditText   ' Item Price
                    '    Dim oEditQuan As SAPbouiCOM.EditText    ' Item Quantity
                    '    Dim oInitQuan As SAPbouiCOM.EditText    ' Item Initial Quantity
                    '    Dim oEditTotal As SAPbouiCOM.EditText   ' Total = Item Price * Item Quantity

                    '    ' Get the items from the matrix
                    '    oEditPrice = colItemPrice.Cells.Item(pVal.Row).Specific
                    '    oEditQuan = colItemQuan.Cells.Item(pVal.Row).Specific
                    '    oInitQuan = colInitQuan.Cells.Item(pVal.Row).Specific
                    '    oEditTotal = colItemTotal.Cells.Item(pVal.Row).Specific

                    '    ' Copy the value to the Initial sum
                    '    Dim tmpInt As Integer
                    '    tmpInt = CInt(oEditQuan.Value)
                    '    oInitQuan.Value = tmpInt

                    '    ' Calc the total column
                    '    Dim tmpTotal As Integer ' temp variable to contain total result
                    '    tmpTotal = CInt(oEditPrice.Value) * CInt(oEditQuan.Value)
                    '    oEditTotal.Value = CInt(tmpTotal)

                    '    ' Calc the document total

                    '    Dim CalcTotal As Double
                    '    Dim i As Integer

                    '    CalcTotal = 0
                    '    ' Iterate all the matrix rows
                    '    For i = 1 To oMatrix.RowCount
                    '        oEditTotal = colItemTotal.Cells.Item(i).Specific
                    '        CalcTotal += oEditTotal.Value
                    '    Next
                    '    oDocTotal.Value = CalcTotal
                    'End If

                End If
            End If
        Catch ex As Exception
            SBOApplication.MessageBox(ex.Message)
        End Try
    End Sub


    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.MenuEvent
        Try


            If (pVal.MenuUID = mnuShipment) And (pVal.BeforeAction = False) Then
                Try
                    oShipmentForm = SBOApplication.Forms.Item(frmShipment)
                    oShipmentForm.Select()
                Catch ex As Exception
                    DrawForm()
                End Try
            End If

            If (pVal.MenuUID = mnuAddRecord) And (pVal.BeforeAction = False) Then
                Dim oForm As SAPbouiCOM.Form
                oForm = SBOApplication.Forms.ActiveForm

                If Not oForm Is Nothing Then
                    If oForm.TypeEx = frmARInvoice _
                            Or oForm.TypeEx = frmARInvoicePayment _
                            Or oForm.TypeEx = frmARCreditMemo _
                            Or oForm.TypeEx = frmARReverseInvoice Then
                        oForm.DataSources.UserDataSources.Item("ds1Territ").ValueEx = ""
                        oForm.DataSources.UserDataSources.Item("ds2Territ").ValueEx = ""
                    End If
                End If
            End If
        Catch ex As Exception
            SBOApplication.MessageBox(ex.Message)
        End Try
    End Sub


    Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBOApplication.FormDataEvent
        Try


            If BusinessObjectInfo.FormUID = frmShipment Then
                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD _
                        And BusinessObjectInfo.BeforeAction = False _
                        And BusinessObjectInfo.ActionSuccess = True Then
                    Dim sVenCode As String
                    Dim oDBDataSource As SAPbouiCOM.DBDataSource
                    oDBDataSource = oShipmentForm.DataSources.DBDataSources.Item("@MAM_OSHP")
                    sVenCode = oDBDataSource.GetValue("U_VenCode", 0).Trim()
                    Dim sCnctCode As String
                    sCnctCode = oDBDataSource.GetValue("U_CnctCode", 0)
                    AddContactPersonsCombo(sVenCode)
                    Dim oComboBox As SAPbouiCOM.ComboBox
                    oComboBox = oShipmentForm.Items.Item(ctlCnctCode).Specific
                    oComboBox.Select(sCnctCode)
                    SetListCondition(oCFLs.Item("5"), "CardCode", sVenCode)
                    SetListCondition(oCFLs.Item("7"), "CardCode", sVenCode)

                    UpdateMatrixColumns()
                    EnableCopyItemsButton()
                End If
            End If

            'If (BusinessObjectInfo.FormTypeEx = frmPurchaseOrders _
            '       Or BusinessObjectInfo.FormTypeEx = frmGoodsReceiptPO _
            '       Or BusinessObjectInfo.FormTypeEx = frmInventoryTransfer) And _
            '            (BusinessObjectInfo.BeforeAction = False) And _
            '            (BusinessObjectInfo.ActionSuccess = True) Then

            '    If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
            '        oDraftBtn.Enable = True
            '    Else
            '        oDraftBtn.Enable = False
            '    End If

            'End If


            If (BusinessObjectInfo.FormTypeEx = frmPurchaseOrders _
                     Or BusinessObjectInfo.FormTypeEx = frmGoodsReceiptPO _
                     Or BusinessObjectInfo.FormTypeEx = frmInventoryTransfer) Then

                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And _
                        BusinessObjectInfo.BeforeAction = False And _
                        nDraftKey <> 0 Then
                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBOApplication.Forms.GetForm(frmDocumentDrafts, 1)
                    Dim RowRemoved As Boolean
                    RowRemoved = False
                    ' Try to remove item for Document Draft list form if is visible
                    If oForm.Visible Then
                        'oForm.Update()
                        Try
                            oForm.Freeze(True)
                            Dim oMatrix As SAPbouiCOM.Matrix
                            oMatrix = oForm.Items.Item("3").Specific
                            Dim oColumn As SAPbouiCOM.Column
                            oColumn = oMatrix.Columns.Item("1")
                            Dim oEditText As SAPbouiCOM.EditText
                            Dim oCell As SAPbouiCOM.Cell
                            Dim nRow As Integer
                            Dim nRowCount As Integer
                            nRowCount = oMatrix.RowCount
                            For nRow = 1 To nRowCount
                                oCell = oColumn.Cells.Item(nRow)
                                oEditText = oCell.Specific
                                If oEditText.Value = nDraftKey Then
                                    oMatrix.DeleteRow(nRow)
                                    'oMatrix.FlushToDataSource()
                                    'RowRemoved = True
                                    Exit For
                                End If
                            Next
                            oForm.Freeze(False)
                            'oForm.DataSources.DBDataSources.Item(0).Query()
                            'oMatrix.LoadFromDataSource()
                        Catch ex As Exception
                            SBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If
                    ' If item is not removed using Document Drafts form 
                    ' try to remove it using business methods API
                    If Not RowRemoved Then
                        Dim oDrafts As SAPbobsCOM.Documents
                        oDrafts = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                        If oDrafts.GetByKey(nDraftKey) Then
                            oDrafts.Remove()
                            RowRemoved = True
                        End If
                    End If
                    nDraftKey = 0
                End If
            End If


            If (BusinessObjectInfo.FormTypeEx = frmPickDetails) Then
                If ((BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) And _
                        (BusinessObjectInfo.BeforeAction = False)) Then

                    Dim nKey As Long
                    nKey = GetKeyFromXML(BusinessObjectInfo.ObjectKey, "PickListParams/AbsoluteEntry")

                    If nKey > 0 Then

                        Dim oPickList As SAPbobsCOM.PickLists
                        oPickList = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists)

                        If oPickList.GetByKey(nKey) Then
                            Dim Rs As SAPbobsCOM.Recordset
                            Rs = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            Dim oPickListLines As SAPbobsCOM.PickLists_Lines
                            oPickListLines = oPickList.Lines

                            Dim nLine As Integer
                            For nLine = 1 To oPickList.Lines.Count
                                oPickListLines.SetCurrentLine(nLine - 1)

                                Dim sItemCode As String
                                Rs.DoQuery("SELECT ItemCode FROM RDR1 WHERE DocEntry = " & oPickListLines.OrderEntry & " AND LineNum = " & oPickListLines.LineNumber)
                                Rs.MoveFirst()
                                If Not Rs.EoF Then
                                    sItemCode = Rs.Fields.Item("ItemCode").Value
                                End If

                                Dim nReleasedQty As Double
                                nReleasedQty = oPickListLines.ReleasedQuantity

                                Dim nRemainQty As Double
                                nRemainQty = nReleasedQty
                                Dim nBinIndex As Integer
                                nBinIndex = 1

                                Rs.DoQuery("SELECT BatchNum, Quantity FROM OIBT WHERE ItemCode = '" & sItemCode & "' AND WhsCode = '01' AND Quantity > 0")
                                Rs.MoveFirst()
                                While Not Rs.EoF
                                    Dim sBinNo As String
                                    Dim nBinQty As Double
                                    sBinNo = Rs.Fields.Item("BatchNum").Value
                                    nBinQty = Rs.Fields.Item("Quantity").Value
                                    If nBinQty > 0 Then
                                        oPickListLines.UserFields.Fields.Item("U_BinNo" & nBinIndex).Value = sBinNo
                                        Try
                                            If nBinQty > nRemainQty Then
                                                oPickListLines.UserFields.Fields.Item("U_BinQty" & nBinIndex).Value = nRemainQty
                                                Exit While
                                            End If
                                            oPickListLines.UserFields.Fields.Item("U_BinQty" & nBinIndex).Value = nBinQty
                                        Catch ex As Exception
                                            SBOApplication.StatusBar.SetText("Error setting Bin Qty (" & nBinIndex & "): " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End Try
                                        nRemainQty = Math.Max(nRemainQty - nBinQty, 0)
                                        nBinIndex = nBinIndex + 1
                                    End If
                                    Rs.MoveNext()
                                End While
                            Next
                            Rs = Nothing

                            Try
                                oPickList.Update()
                            Catch ex As Exception
                                SBOApplication.StatusBar.SetText("Error updating pick list bins. Reason: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try

                            System.GC.Collect()

                        End If

                    End If
                End If
            End If

            If (BusinessObjectInfo.FormTypeEx = frmARInvoice _
                     Or BusinessObjectInfo.FormTypeEx = frmARInvoicePayment _
                     Or BusinessObjectInfo.FormTypeEx = frmARCreditMemo _
                     Or BusinessObjectInfo.FormTypeEx = frmARReverseInvoice) Then


                If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD _
                        And BusinessObjectInfo.BeforeAction = False Then

                    Dim oForm As SAPbouiCOM.Form
                    Try

                        oForm = SBOApplication.Forms.ActiveForm() 'SBOApplication.Forms.GetForm(BusinessObjectInfo.FormTypeEx, 1)
                    Catch ex As Exception
                        'MsgBox(ex.Message)
                    End Try
                    Dim oEditText As SAPbouiCOM.EditText
                    Try
                        oEditText = oForm.Items.Item("bndTerrit").Specific
                        Dim sTerritId As String
                        Dim sTerritDesc As String
                        sTerritId = oEditText.Value
                        If sTerritId <> "" Then
                            Dim oRs As SAPbobsCOM.Recordset
                            oRs = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRs.DoQuery("SELECT descript FROM OTER WHERE territryID = " & sTerritId)
                            oRs.MoveFirst()
                            If Not oRs.EoF Then
                                sTerritDesc = oRs.Fields.Item("descript").Value
                            Else
                                sTerritDesc = ""
                            End If
                        Else
                            sTerritDesc = ""
                        End If
                        oForm.DataSources.UserDataSources.Item("ds1Territ").ValueEx = sTerritDesc
                        oForm.DataSources.UserDataSources.Item("ds2Territ").ValueEx = sTerritId
                    Catch ex As Exception
                        ' MsgBox(ex.Message)
                    End Try
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SBO_Application_StatusBarEvent(ByVal Text As String, ByVal MessageType As SAPbouiCOM.BoStatusBarMessageType) Handles SBOApplication.StatusBarEvent
        Try

        
            If AddStarted = True And MessageType = SAPbouiCOM.BoStatusBarMessageType.smt_Error Then
                AddStarted = False
            End If
        Catch ex As Exception
            SBOApplication.MessageBox(ex.Message)
        End Try
    End Sub

    Private Sub SBOApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBOApplication.AppEvent

        If EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Then

            System.Windows.Forms.Application.Exit()

        ElseIf EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then

            System.Windows.Forms.Application.Exit()

        End If

    End Sub

End Class
