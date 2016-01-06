Public Class frmUDO
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents grpConnection As System.Windows.Forms.GroupBox
    Friend WithEvents lblCompany As System.Windows.Forms.Label
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents cmbCompany As System.Windows.Forms.ComboBox
    Friend WithEvents btnGetCompanyList As System.Windows.Forms.Button
    Friend WithEvents txtPass As System.Windows.Forms.TextBox
    Friend WithEvents txtUser As System.Windows.Forms.TextBox
    Friend WithEvents lblPass As System.Windows.Forms.Label
    Friend WithEvents lblUser As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnConnect As System.Windows.Forms.Button
    Friend WithEvents lblIntro As System.Windows.Forms.Label
    Friend WithEvents grpUDO As System.Windows.Forms.GroupBox
    Friend WithEvents chkUDOAfter As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnAddFields As System.Windows.Forms.Button
    Friend WithEvents btnAddUDO As System.Windows.Forms.Button
    Friend WithEvents txtLog As System.Windows.Forms.TextBox
    Friend WithEvents btnAddUserObjects As System.Windows.Forms.Button
    Friend WithEvents chkUseTrusted As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.grpConnection = New System.Windows.Forms.GroupBox
        Me.chkUseTrusted = New System.Windows.Forms.CheckBox
        Me.lblCompany = New System.Windows.Forms.Label
        Me.txtServer = New System.Windows.Forms.TextBox
        Me.cmbCompany = New System.Windows.Forms.ComboBox
        Me.btnGetCompanyList = New System.Windows.Forms.Button
        Me.txtPass = New System.Windows.Forms.TextBox
        Me.txtUser = New System.Windows.Forms.TextBox
        Me.lblPass = New System.Windows.Forms.Label
        Me.lblUser = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnConnect = New System.Windows.Forms.Button
        Me.lblIntro = New System.Windows.Forms.Label
        Me.grpUDO = New System.Windows.Forms.GroupBox
        Me.btnAddUserObjects = New System.Windows.Forms.Button
        Me.txtLog = New System.Windows.Forms.TextBox
        Me.btnAddFields = New System.Windows.Forms.Button
        Me.chkUDOAfter = New System.Windows.Forms.CheckedListBox
        Me.btnAdd = New System.Windows.Forms.Button
        Me.btnAddUDO = New System.Windows.Forms.Button
        Me.grpConnection.SuspendLayout()
        Me.grpUDO.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpConnection
        '
        Me.grpConnection.Controls.Add(Me.chkUseTrusted)
        Me.grpConnection.Controls.Add(Me.lblCompany)
        Me.grpConnection.Controls.Add(Me.txtServer)
        Me.grpConnection.Controls.Add(Me.cmbCompany)
        Me.grpConnection.Controls.Add(Me.btnGetCompanyList)
        Me.grpConnection.Controls.Add(Me.txtPass)
        Me.grpConnection.Controls.Add(Me.txtUser)
        Me.grpConnection.Controls.Add(Me.lblPass)
        Me.grpConnection.Controls.Add(Me.lblUser)
        Me.grpConnection.Controls.Add(Me.Label1)
        Me.grpConnection.Controls.Add(Me.btnConnect)
        Me.grpConnection.Location = New System.Drawing.Point(8, 56)
        Me.grpConnection.Name = "grpConnection"
        Me.grpConnection.Size = New System.Drawing.Size(544, 120)
        Me.grpConnection.TabIndex = 18
        Me.grpConnection.TabStop = False
        Me.grpConnection.Text = "Connection"
        '
        'chkUseTrusted
        '
        Me.chkUseTrusted.Checked = True
        Me.chkUseTrusted.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkUseTrusted.Location = New System.Drawing.Point(360, 64)
        Me.chkUseTrusted.Name = "chkUseTrusted"
        Me.chkUseTrusted.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkUseTrusted.Size = New System.Drawing.Size(176, 24)
        Me.chkUseTrusted.TabIndex = 29
        Me.chkUseTrusted.Text = "Use Trusted Connection"
        '
        'lblCompany
        '
        Me.lblCompany.Location = New System.Drawing.Point(360, 32)
        Me.lblCompany.Name = "lblCompany"
        Me.lblCompany.Size = New System.Drawing.Size(56, 16)
        Me.lblCompany.TabIndex = 27
        Me.lblCompany.Text = "Company"
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(112, 32)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.TabIndex = 18
        Me.txtServer.Text = "(local)"
        '
        'cmbCompany
        '
        Me.cmbCompany.Location = New System.Drawing.Point(416, 32)
        Me.cmbCompany.Name = "cmbCompany"
        Me.cmbCompany.Size = New System.Drawing.Size(121, 21)
        Me.cmbCompany.TabIndex = 21
        '
        'btnGetCompanyList
        '
        Me.btnGetCompanyList.Location = New System.Drawing.Point(224, 32)
        Me.btnGetCompanyList.Name = "btnGetCompanyList"
        Me.btnGetCompanyList.Size = New System.Drawing.Size(128, 23)
        Me.btnGetCompanyList.TabIndex = 19
        Me.btnGetCompanyList.Text = "Get Company List"
        '
        'txtPass
        '
        Me.txtPass.Location = New System.Drawing.Point(136, 88)
        Me.txtPass.Name = "txtPass"
        Me.txtPass.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtPass.TabIndex = 24
        Me.txtPass.Text = "manager"
        '
        'txtUser
        '
        Me.txtUser.Location = New System.Drawing.Point(136, 64)
        Me.txtUser.Name = "txtUser"
        Me.txtUser.TabIndex = 22
        Me.txtUser.Text = "manager"
        '
        'lblPass
        '
        Me.lblPass.Location = New System.Drawing.Point(16, 88)
        Me.lblPass.Name = "lblPass"
        Me.lblPass.Size = New System.Drawing.Size(112, 16)
        Me.lblPass.TabIndex = 26
        Me.lblPass.Text = "Database Password"
        '
        'lblUser
        '
        Me.lblUser.Location = New System.Drawing.Point(16, 64)
        Me.lblUser.Name = "lblUser"
        Me.lblUser.Size = New System.Drawing.Size(112, 16)
        Me.lblUser.TabIndex = 23
        Me.lblUser.Text = "Database Username"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 16)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "Database Server"
        '
        'btnConnect
        '
        Me.btnConnect.Location = New System.Drawing.Point(248, 88)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.TabIndex = 25
        Me.btnConnect.Text = "Connect"
        '
        'lblIntro
        '
        Me.lblIntro.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.lblIntro.Location = New System.Drawing.Point(16, 8)
        Me.lblIntro.Name = "lblIntro"
        Me.lblIntro.Size = New System.Drawing.Size(416, 40)
        Me.lblIntro.TabIndex = 19
        Me.lblIntro.Text = "This application creates the user tables, adds user fields and registers UDO obje" & _
        "ct for MAMCO SAP Business One."
        '
        'grpUDO
        '
        Me.grpUDO.Controls.Add(Me.btnAddUserObjects)
        Me.grpUDO.Controls.Add(Me.txtLog)
        Me.grpUDO.Controls.Add(Me.btnAddFields)
        Me.grpUDO.Controls.Add(Me.chkUDOAfter)
        Me.grpUDO.Controls.Add(Me.btnAdd)
        Me.grpUDO.Controls.Add(Me.btnAddUDO)
        Me.grpUDO.Location = New System.Drawing.Point(8, 192)
        Me.grpUDO.Name = "grpUDO"
        Me.grpUDO.Size = New System.Drawing.Size(544, 264)
        Me.grpUDO.TabIndex = 22
        Me.grpUDO.TabStop = False
        Me.grpUDO.Text = "User Defined Tables,Fields and Objects"
        '
        'btnAddUserObjects
        '
        Me.btnAddUserObjects.Enabled = False
        Me.btnAddUserObjects.Location = New System.Drawing.Point(368, 12)
        Me.btnAddUserObjects.Name = "btnAddUserObjects"
        Me.btnAddUserObjects.Size = New System.Drawing.Size(168, 24)
        Me.btnAddUserObjects.TabIndex = 25
        Me.btnAddUserObjects.Text = "Add User Tables and Fields"
        '
        'txtLog
        '
        Me.txtLog.AutoSize = False
        Me.txtLog.Location = New System.Drawing.Point(8, 40)
        Me.txtLog.Multiline = True
        Me.txtLog.Name = "txtLog"
        Me.txtLog.Size = New System.Drawing.Size(528, 216)
        Me.txtLog.TabIndex = 24
        Me.txtLog.Text = ""
        '
        'btnAddFields
        '
        Me.btnAddFields.Location = New System.Drawing.Point(8, 72)
        Me.btnAddFields.Name = "btnAddFields"
        Me.btnAddFields.Size = New System.Drawing.Size(80, 23)
        Me.btnAddFields.TabIndex = 4
        Me.btnAddFields.Text = "Add Fields"
        Me.btnAddFields.Visible = False
        '
        'chkUDOAfter
        '
        Me.chkUDOAfter.Items.AddRange(New Object() {"Shipment Document (OSHP)", "... Vendor Code (VenCode)", "... Vendor Name (VenName)", "... Contact Person (CnctCode)", "... Vendor Ref. No. (VenRefNo)", "... Container Number (Containr)", "... Proforma Number (ProForma)", "... Seal Number (SealNum)", "... Shipping Agent (Agent)", "... Vessel Name (Vessel)", "... Due Date (DueDate)", "... Arrival Date (ArvDate)", "... Remarks (Comments)", "Shipment Lines (SHP1)", "... Base Document Reference (BaseRef)", "... Base Document Key (BaseEntry)", "... Base Row (BaseLine)", "... Row Status (LineStatus)", "... Item No. (ItemCode)", "... Item Description (ItemName)", "... Quantity (Quantity)", "... Shipped Quantity (ShipQty)", "... Price (Price)", "... Price Currency (Currency)", "... Unit of measure (UnitMsr)", "... Carton unit of measure (CrtUnit)"})
        Me.chkUDOAfter.Location = New System.Drawing.Point(104, 32)
        Me.chkUDOAfter.Name = "chkUDOAfter"
        Me.chkUDOAfter.Size = New System.Drawing.Size(248, 214)
        Me.chkUDOAfter.TabIndex = 2
        Me.chkUDOAfter.Visible = False
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(8, 32)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(80, 23)
        Me.btnAdd.TabIndex = 0
        Me.btnAdd.Text = "Add Tables"
        Me.btnAdd.Visible = False
        '
        'btnAddUDO
        '
        Me.btnAddUDO.Location = New System.Drawing.Point(8, 112)
        Me.btnAddUDO.Name = "btnAddUDO"
        Me.btnAddUDO.Size = New System.Drawing.Size(80, 23)
        Me.btnAddUDO.TabIndex = 6
        Me.btnAddUDO.Text = "Add UDO"
        Me.btnAddUDO.Visible = False
        '
        'frmUDO
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(560, 469)
        Me.Controls.Add(Me.grpUDO)
        Me.Controls.Add(Me.lblIntro)
        Me.Controls.Add(Me.grpConnection)
        Me.Name = "frmUDO"
        Me.Text = "UDO registration"
        Me.grpConnection.ResumeLayout(False)
        Me.grpUDO.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Public oCompany As SAPbobsCOM.Company

    ' Error handling variables
    Public sErrMsg As String
    Public lErrCode As Integer
    Public lRetCode As Integer

    Private Sub frmUDO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        oCompany = New SAPbobsCOM.Company

        '// once the Server property of the Company is set
        '// we may query for a list of companies to choos from
        '// this method returns a Recordset object



    End Sub

    Private Sub btnGetCompanyList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetCompanyList.Click

        oCompany.Server = txtServer.Text
        oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
        If chkUseTrusted.Checked Then
            oCompany.UseTrusted = True
        Else
            oCompany.DbUserName = txtUser.Text
            oCompany.DbPassword = txtPass.Text
            oCompany.UseTrusted = False
        End If

        Dim oRecordSet As SAPbobsCOM.Recordset

        oRecordSet = oCompany.GetCompanyList

        '// Use GetLastError method directly after a function
        '// which doesn't have a return code
        '// you may also use the On Error GoTo.
        '// functions with no return codes throws exceptions

        oCompany.GetLastError(lErrCode, sErrMsg)

        If lErrCode <> 0 Then
            MsgBox(sErrMsg)
        Else
            '// Load the available company DB names to the combo box
            '// the returned Recordset containds 4 fields:
            '// dbName - represents the database name
            '// cmpName - represents the company name
            '// versStr - represents the version number of the company database
            '// dbUser - represents the database owner
            '// we are interested in the first filed (mandatory property)

            '// Go through the Recordset and extract the dbname
            Do Until oRecordSet.EoF = True
                '// add the value of the first field of the Recordset
                cmbCompany.Items.Add(oRecordSet.Fields.Item(0).Value)
                '// move the record pointer to the next row
                oRecordSet.MoveNext()
            Loop
            If cmbCompany.Items.Count > 0 Then
                cmbCompany.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub btnConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect.Click

        ' Set the connection parameters
        oCompany.CompanyDB = cmbCompany.Text
        oCompany.UserName = txtUser.Text
        oCompany.Password = txtPass.Text

        lErrCode = oCompany.Connect

        If lErrCode <> 0 Then
            oCompany.GetLastError(lErrCode, sErrMsg)
            txtLog.AppendText(sErrMsg & vbCrLf)
        Else
            grpConnection.Enabled = False
            'grpUDO.Enabled = True
            btnAddUserObjects.Enabled = True
            txtLog.AppendText("Connected to " & oCompany.CompanyName & vbCrLf)
        End If

    End Sub


    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        AddUserTable("MAM_OSHP", "Shipment Document", SAPbobsCOM.BoUTBTableType.bott_Document)
        chkUDOAfter.SetItemChecked(0, True)
        AddUserTable("MAM_SHP1", "Shipment Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        chkUDOAfter.SetItemChecked(13, True)
    End Sub

    Private Function AddOSHPFields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        Dim nPos As Integer

        '************************************
        ' Adding "Vendor Code" field
        '************************************
        '// Setting the Field's properties
        nPos = 1
        oUserFieldsMD.TableName = "@MAM_OSHP"
        oUserFieldsMD.Name = "VenCode"
        oUserFieldsMD.Description = "Vendor Code"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Vendor Name" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_OSHP"
        oUserFieldsMD.Name = "VenName"
        oUserFieldsMD.Description = "Vendor Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Vendor Ref. No." field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_OSHP"
        oUserFieldsMD.Name = "VenRefNo"
        oUserFieldsMD.Description = "Vendor Ref. No."
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Contact Name" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_OSHP"
        oUserFieldsMD.Name = "CnctCode"
        oUserFieldsMD.Description = "Contact person"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Numeric
        oUserFieldsMD.EditSize = 11

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Container number" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_OSHP"
        oUserFieldsMD.Name = "Containr"
        oUserFieldsMD.Description = "Container number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Pro forma number" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_OSHP"
        oUserFieldsMD.Name = "ProForma"
        oUserFieldsMD.Description = "Pro forma number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Seal number" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_OSHP"
        oUserFieldsMD.Name = "SealNum"
        oUserFieldsMD.Description = "Seal number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Vessel name" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_OSHP"
        oUserFieldsMD.Name = "Agent"
        oUserFieldsMD.Description = "Shipping agent"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 50

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Vessel name" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_OSHP"
        oUserFieldsMD.Name = "Vessel"
        oUserFieldsMD.Description = "Vessel name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 50

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Due Date" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_OSHP"
        oUserFieldsMD.Name = "DueDate"
        oUserFieldsMD.Description = "Due Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Arrival Date" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_OSHP"
        oUserFieldsMD.Name = "ArvDate"
        oUserFieldsMD.Description = "Arrival Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Remarks" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_OSHP"
        oUserFieldsMD.Name = "Comments"
        oUserFieldsMD.Description = "Remarks"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 254

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        GC.Collect() 'Release the handle to the User Fields
    End Function

    Private Function AddSHP1Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        Dim nPos As Integer

        '************************************
        ' Adding "Base Document Reference" field
        '************************************
        '// Setting the Field's properties

        nPos = 14
        oUserFieldsMD.TableName = "@MAM_SHP1"
        oUserFieldsMD.Name = "BaseRef"
        oUserFieldsMD.Description = "Base Document Reference"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 16

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Base Document Key" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_SHP1"
        oUserFieldsMD.Name = "BaseEntry"
        oUserFieldsMD.Description = "Base Document Key"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Numeric
        oUserFieldsMD.EditSize = 11

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Base Row" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_SHP1"
        oUserFieldsMD.Name = "BaseLine"
        oUserFieldsMD.Description = "Base Row"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Numeric
        oUserFieldsMD.EditSize = 11

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Row Status (O/C)" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_SHP1"
        oUserFieldsMD.Name = "LnStatus"
        oUserFieldsMD.Description = "Row Status (O/C)"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 1

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Item No." field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_SHP1"
        oUserFieldsMD.Name = "ItemCode"
        oUserFieldsMD.Description = "Item No."
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Item Description" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_SHP1"
        oUserFieldsMD.Name = "ItemName"
        oUserFieldsMD.Description = "Item Description"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_SHP1"
        oUserFieldsMD.Name = "Price"
        oUserFieldsMD.Description = "Price"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Quantity" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_SHP1"
        oUserFieldsMD.Name = "Quantity"
        oUserFieldsMD.Description = "Quantity"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Quantity
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Shipped Quantity" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_SHP1"
        oUserFieldsMD.Name = "ShipQty"
        oUserFieldsMD.Description = "Shipped Quantity"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Quantity
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Currency" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_SHP1"
        oUserFieldsMD.Name = "Currency"
        oUserFieldsMD.Description = "Currency"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 3

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Unit of measure" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_SHP1"
        oUserFieldsMD.Name = "UnitMsr"
        oUserFieldsMD.Description = "Unit of measure"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        '************************************
        ' Adding "Cartoon unit of measure" field
        '************************************
        '// Setting the Field's properties

        nPos = nPos + 1
        oUserFieldsMD.TableName = "@MAM_SHP1"
        oUserFieldsMD.Name = "CrtUnit"
        oUserFieldsMD.Description = "Cartoon unit of measure"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            MsgBox(sErrMsg)
        Else
            chkUDOAfter.SetItemChecked(nPos, True)
            MsgBox("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table")
        End If

        GC.Collect() 'Release the handle to the User Fields
    End Function

    Private Function TableExist(ByVal TableName As String) As Boolean
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Dim bool As Boolean
        oUserTablesMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)

        bool = oUserTablesMD.GetByKey(TableName)
        Return (bool)
    End Function

    Private Function FieldExist()

    End Function

    Private Sub btnAddFields_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddFields.Click
        AddOSHPFields()
        AddSHP1Fields()
    End Sub

    Private Sub AddUDO()

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD

        oUserObjectMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

        oUserObjectMD.Code = "MAM_SHP"
        oUserObjectMD.Name = "Shipment Document"
        oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document
        oUserObjectMD.TableName = "MAM_OSHP"
        oUserObjectMD.ChildTables.TableName = "MAM_SHP1"
        oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO

        lRetCode = oUserObjectMD.Add()

        '// check for errors in the process
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
                chkUDOAfter.SetItemChecked(12, True)
            Else
                oCompany.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            MsgBox("UDO: " & oUserObjectMD.Name & " was added successfully")
            chkUDOAfter.SetItemChecked(14, True)
        End If

        oUserObjectMD = Nothing

        GC.Collect() 'Release the handle to the table
    End Sub

    Private Sub btnAddUDO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddUDO.Click
        AddUDO()
    End Sub

    Private Sub AddUserFields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        'Dim oDocuments As SAPbobsCOM.Documents
        'oDocuments = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)


        '************************************
        ' Marketing Title Fields
        '************************************

        '************************************
        ' Adding "Agent" field
        '************************************
        '// Setting the Field's properties
        oUserFieldsMD.TableName = "OPOR"
        oUserFieldsMD.Name = "Agent"
        oUserFieldsMD.Description = "Shipping Agent"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 30

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            txtLog.AppendText("[FAIL] Field '" & oUserFieldsMD.Name & "' was not added to '" & oUserFieldsMD.TableName & "' table" & vbCrLf)
            txtLog.AppendText("       Reason: " & sErrMsg & vbCrLf)
        Else
            txtLog.AppendText("[OK]   Field '" & oUserFieldsMD.Name & "' was added successfuly to '" & oUserFieldsMD.TableName & "' table" & vbCrLf)
        End If


        '************************************
        ' Adding "Vessel" field
        '************************************
        '// Setting the Field's properties
        oUserFieldsMD.TableName = "OPOR"
        oUserFieldsMD.Name = "Vessel"
        oUserFieldsMD.Description = "Vessel"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 30

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            txtLog.AppendText("[FAIL] Field '" & oUserFieldsMD.Name & "' was not added to '" & oUserFieldsMD.TableName & "' table" & vbCrLf)
            txtLog.AppendText("       Reason: " & sErrMsg & vbCrLf)
        Else
            txtLog.AppendText("[OK]   Field '" & oUserFieldsMD.Name & "' was added successfuly to '" & oUserFieldsMD.TableName & "' table" & vbCrLf)
        End If


        '************************************
        ' Adding "Seal number" field
        '************************************
        '// Setting the Field's properties
        oUserFieldsMD.TableName = "OPOR"
        oUserFieldsMD.Name = "SealNum"
        oUserFieldsMD.Description = "Seal number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            txtLog.AppendText("[FAIL] Field '" & oUserFieldsMD.Name & "' was not added to '" & oUserFieldsMD.TableName & "' table")
            txtLog.AppendText("       Reason: " & sErrMsg)
        Else
            txtLog.AppendText("[OK]   Field '" & oUserFieldsMD.Name & "' was added successfuly to '" & oUserFieldsMD.TableName & "' table")
        End If

        '************************************
        ' Adding "Container number" field
        '************************************
        '// Setting the Field's properties
        oUserFieldsMD.TableName = "OPOR"
        oUserFieldsMD.Name = "Containr"
        oUserFieldsMD.Description = "Container number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            txtLog.AppendText("[FAIL] Field '" & oUserFieldsMD.Name & "' was not added to '" & oUserFieldsMD.TableName & "' table")
            txtLog.AppendText("       Reason: " & sErrMsg)
        Else
            txtLog.AppendText("[OK]   Field '" & oUserFieldsMD.Name & "' was added successfuly to '" & oUserFieldsMD.TableName & "' table")
        End If

        '************************************
        ' Pick List Rows Fields
        '************************************

        '************************************
        ' Adding "Bin No 1" field
        '************************************
        '// Setting the Field's properties
        oUserFieldsMD.TableName = "PKL1"
        oUserFieldsMD.Name = "BinNo1"
        oUserFieldsMD.Description = "Bin No (1)"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            txtLog.AppendText("[FAIL] Field '" & oUserFieldsMD.Name & "' was not added to '" & oUserFieldsMD.TableName & "' table")
            txtLog.AppendText("       Reason: " & sErrMsg)
        Else
            txtLog.AppendText("[OK]   Field '" & oUserFieldsMD.Name & "' was added successfuly to '" & oUserFieldsMD.TableName & "' table")
        End If

        '************************************
        ' Adding "Bin No 2" field
        '************************************
        '// Setting the Field's properties
        oUserFieldsMD.TableName = "PKL1"
        oUserFieldsMD.Name = "BinNo2"
        oUserFieldsMD.Description = "Bin No (2)"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            txtLog.AppendText("[FAIL] Field '" & oUserFieldsMD.Name & "' was not added to '" & oUserFieldsMD.TableName & "' table")
            txtLog.AppendText("       Reason: " & sErrMsg)
        Else
            txtLog.AppendText("[OK]   Field '" & oUserFieldsMD.Name & "' was added successfuly to '" & oUserFieldsMD.TableName & "' table")
        End If

        '************************************
        ' Adding "Bin Qty 1" field
        '************************************
        '// Setting the Field's properties
        oUserFieldsMD.TableName = "PKL1"
        oUserFieldsMD.Name = "BinQty1"
        oUserFieldsMD.Description = "Bin Qty (1)"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Numeric
        oUserFieldsMD.EditSize = 11

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            txtLog.AppendText("[FAIL] Field '" & oUserFieldsMD.Name & "' was not added to '" & oUserFieldsMD.TableName & "' table")
            txtLog.AppendText("       Reason: " & sErrMsg)
        Else
            txtLog.AppendText("[OK]   Field '" & oUserFieldsMD.Name & "' was added successfuly to '" & oUserFieldsMD.TableName & "' table")
        End If

        '************************************
        ' Adding "Bin Qty 1" field
        '************************************
        '// Setting the Field's properties
        oUserFieldsMD.TableName = "PKL1"
        oUserFieldsMD.Name = "BinQty2"
        oUserFieldsMD.Description = "Bin Qty (2)"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Numeric
        oUserFieldsMD.EditSize = 11

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            txtLog.AppendText("[FAIL] Field '" & oUserFieldsMD.Name & "' was not added to '" & oUserFieldsMD.TableName & "' table")
            txtLog.AppendText("       Reason: " & sErrMsg)
        Else
            txtLog.AppendText("[OK]   Field '" & oUserFieldsMD.Name & "' was added successfuly to '" & oUserFieldsMD.TableName & "' table")
        End If

        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub AddUserTable(ByVal Name As String, ByVal Description As String, _
        ByVal Type As SAPbobsCOM.BoUTBTableType)
        '//****************************************************************************
        '// The UserTablesMD represents a meta-data object which allows us
        '// to add\remove tables, change a table name etc.
        '//****************************************************************************

        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD

        '//****************************************************************************
        '// In any meta-data operation there should be no other object "alive"
        '// but the meta-data object, otherwise the operation will fail.
        '// This restriction is intended to prevent a collisions
        '//****************************************************************************

        '// the meta-data object needs to be initialized with a
        '// regular UserTables object
        oUserTablesMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)

        '//**************************************************
        '// when adding user tables or fields to the SBO DB
        '// use a prefix identifying your partner name space
        '// this will prevent collisions between different
        '// partners add-ons
        '//
        '// SAP's name space prefix is "BE_"
        '//**************************************************		

        '// set the table parameters
        oUserTablesMD.TableName = Name
        oUserTablesMD.TableDescription = Description
        oUserTablesMD.TableType = Type

        '// Add the table
        '// This action add an empty table with 2 default fields
        '// 'Code' and 'Name' which serve as the key
        '// in order to add your own User Fields
        '// see the AddUserFields.frm in this project
        '// a privat, user defined, key may be added
        '// see AddPrivateKey.frm in this project

        lRetCode = oUserTablesMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany.GetLastError(lRetCode, sErrMsg)
                txtLog.AppendText("[FAIL] Table '" & oUserTablesMD.TableName & "' was not added" & vbCrLf)
                txtLog.AppendText("       Reason: " & sErrMsg & vbCrLf)
            End If
        Else
            txtLog.AppendText("[OK]   Table '" & oUserTablesMD.TableName & "' was added successfully" & vbCrLf)
        End If

        oUserTablesMD = Nothing

        GC.Collect() 'Release the handle to the table
    End Sub

    Private Sub AddUserField(ByVal TableName As String, ByVal FieldName As String, ByVal Description As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal EditSize As Integer)

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '// Setting the Field's properties
        oUserFieldsMD.TableName = TableName
        oUserFieldsMD.Name = FieldName
        oUserFieldsMD.Description = Description
        oUserFieldsMD.Type = FieldType
        oUserFieldsMD.EditSize = EditSize

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            txtLog.AppendText("[FAIL] Field '" & oUserFieldsMD.Name & "' was not added to '" & oUserFieldsMD.TableName & "' table" & vbCrLf)
            txtLog.AppendText("       Reason: " & sErrMsg & vbCrLf)
        Else
            txtLog.AppendText("[OK]   Field '" & oUserFieldsMD.Name & "' was added successfully to '" & oUserFieldsMD.TableName & "' table" & vbCrLf)
        End If

        oUserFieldsMD = Nothing

        GC.Collect() 'Release the handle to the table

    End Sub


    Private Sub AddUserField(ByVal TableName As String, ByVal FieldName As String, ByVal Description As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal EditSize As Integer, ByVal ValidValueNames() As String, ByVal ValidValueDescs() As String)

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '// Setting the Field's properties
        oUserFieldsMD.TableName = TableName
        oUserFieldsMD.Name = FieldName
        oUserFieldsMD.Description = Description
        oUserFieldsMD.Type = FieldType
        oUserFieldsMD.EditSize = EditSize

        Dim nIndex As Integer
        For nIndex = 1 To UBound(ValidValueNames)
            oUserFieldsMD.ValidValues.Value = ValidValueNames(nIndex)
            oUserFieldsMD.ValidValues.Description = ValidValueDescs(nIndex)
            oUserFieldsMD.ValidValues.Add()
        Next
        oUserFieldsMD.DefaultValue = ValidValueNames(1)

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            txtLog.AppendText("[FAIL] Field '" & oUserFieldsMD.Name & "' was not added to '" & oUserFieldsMD.TableName & "' table" & vbCrLf)
            txtLog.AppendText("       Reason: " & sErrMsg & vbCrLf)
        Else
            txtLog.AppendText("[OK]   Field '" & oUserFieldsMD.Name & "' was added successfully to '" & oUserFieldsMD.TableName & "' table" & vbCrLf)
        End If

        oUserFieldsMD = Nothing

        GC.Collect() 'Release the handle to the table

    End Sub


    Private Sub AddUserField(ByVal TableName As String, ByVal FieldName As String, ByVal Description As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal EditSize As Integer, ByVal LinkedTable As String)

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '// Setting the Field's properties
        oUserFieldsMD.TableName = TableName
        oUserFieldsMD.Name = FieldName
        oUserFieldsMD.Description = Description
        oUserFieldsMD.Type = FieldType
        oUserFieldsMD.EditSize = EditSize
        oUserFieldsMD.LinkedTable = LinkedTable

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            txtLog.AppendText("[FAIL] Field '" & oUserFieldsMD.Name & "' was not added to '" & oUserFieldsMD.TableName & "' table" & vbCrLf)
            txtLog.AppendText("       Reason: " & sErrMsg & vbCrLf)
        Else
            txtLog.AppendText("[OK]   Field '" & oUserFieldsMD.Name & "' was added successfully to '" & oUserFieldsMD.TableName & "' table" & vbCrLf)
        End If

        oUserFieldsMD = Nothing

        GC.Collect() 'Release the handle to the table

    End Sub


    Private Sub AddUserField(ByVal TableName As String, ByVal FieldName As String, ByVal Description As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal FieldSubType As SAPbobsCOM.BoFldSubTypes)

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '// Setting the Field's properties
        oUserFieldsMD.TableName = TableName
        oUserFieldsMD.Name = FieldName
        oUserFieldsMD.Description = Description
        oUserFieldsMD.Type = FieldType
        oUserFieldsMD.SubType = FieldSubType

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            oCompany.GetLastError(lRetCode, sErrMsg)
            txtLog.AppendText("[FAIL] Field '" & oUserFieldsMD.Name & "' was not added to '" & oUserFieldsMD.TableName & "' table" & vbCrLf)
            txtLog.AppendText("       Reason: " & sErrMsg & vbCrLf)
        Else
            txtLog.AppendText("[OK]   Field '" & oUserFieldsMD.Name & "' was added successfully to '" & oUserFieldsMD.TableName & "' table" & vbCrLf)
        End If

        oUserFieldsMD = Nothing

        GC.Collect() 'Release the handle to the table

    End Sub


    Private Sub btnAddUserObjects_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddUserObjects.Click

        btnAddUserObjects.Enabled = False

        ' User Tables
        AddUserTable("MAM_IBRAND", "Item Brand", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        AddUserTable("MAM_ISBRND", "Item Sub Brand", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        AddUserTable("MAM_IGRP", "Item Group", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        AddUserTable("MAM_ISGRP", "Item Sub Group", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        AddUserTable("MAM_ICAT", "Item Category", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        AddUserTable("MAM_IDSGN", "Item Design", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        AddUserTable("MAM_IDIV", "Item Division", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        AddUserTable("MAM_ITYPE", "Item Type", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        AddUserTable("MAM_LC", "Letter of Credit", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        AddUserTable("MAM_DRV", "Driver", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        AddUserTable("MAM_TR", "Territories", SAPbobsCOM.BoUTBTableType.bott_NoObject)

        AddUserField("OITM", "Brand", "Brand", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, "MAM_IBRAND")
        AddUserField("OITM", "SubBrand", "Sub Brand", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, "MAM_ISBRND")
        AddUserField("OITM", "Group", "Group", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, "MAM_IGRP")
        AddUserField("OITM", "SubGroup", "Sub Group", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, "MAM_ISGRP")
        AddUserField("OITM", "Category", "Category", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, "MAM_ICAT")
        AddUserField("OITM", "Design", "Design", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, "MAM_IDSGN")
        AddUserField("OITM", "Division", "Division", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, "MAM_IDIV")
        AddUserField("OITM", "Type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, "MAM_ITYPE")

        ' Marketing Title Fields
        AddUserField("OPOR", "Agent", "Shipping Agent", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        AddUserField("OPOR", "Vessel", "Vessel", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        AddUserField("OPOR", "SealNum", "Seal number", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
        AddUserField("OPOR", "Containr", "Container number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        AddUserField("OPOR", "LcCode", "Letter of Credit", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, "MAM_LC")
        AddUserField("OPOR", "Driver", "Driver", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, "MAM_DRV")
        'AddUserField("OPOR", "TerritId", "Territory", SAPbobsCOM.BoFieldTypes.db_Alpha, 8, "MAM_TR")
        AddUserField("OPOR", "Territ", "Territory", SAPbobsCOM.BoFieldTypes.db_Numeric, 10)

        ' Pick List Rows Fields
        AddUserField("PKL1", "BinNo1", "Bin No (1)", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
        AddUserField("PKL1", "BinQty1", "Bin Qty (1)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        AddUserField("PKL1", "BinNo2", "Bin No (2)", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
        AddUserField("PKL1", "BinQty2", "Bin Qty (2)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity)
        AddUserField("PKL1", "BinNo3", "Bin No (3)", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
        AddUserField("PKL1", "BinQty3", "Bin Qty (3)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity)

        ' Letter of Credit Fields
        Dim LcStatusNames(2) As String
        LcStatusNames(1) = "01"
        LcStatusNames(2) = "02"
        Dim LcStatusDescs(2) As String
        LcStatusDescs(1) = "Outstanding"
        LcStatusDescs(2) = "Closed"
        AddUserField("@MAM_LC", "Status", "LC Status", SAPbobsCOM.BoFieldTypes.db_Numeric, 2, LcStatusNames, LcStatusDescs)
        AddUserField("@MAM_LC", "Amount", "LC Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price)
        AddUserField("@MAM_LC", "Date", "LC Due Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None)

        txtLog.AppendText("Done." & vbCrLf)

        'AddMarketingFields("OPDF")
        'AddMarketingFields("ORPD")
        'AddMarketingFields("OPCH")
        'AddMarketingFields("ODRF")
        btnAddUserObjects.Enabled = True

    End Sub
End Class
