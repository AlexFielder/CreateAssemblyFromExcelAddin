<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FileSelectionForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.SearchResultsListBox = New System.Windows.Forms.ListBox()
        Me.m_SearchingForLabel = New System.Windows.Forms.Label()
        Me.BtnDone = New System.Windows.Forms.Button()
        Me.BtnNoFile = New System.Windows.Forms.Button()
        Me.m_itemsCountLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'SearchResultsListBox
        '
        Me.SearchResultsListBox.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SearchResultsListBox.Location = New System.Drawing.Point(12, 42)
        Me.SearchResultsListBox.Name = "SearchResultsListBox"
        Me.SearchResultsListBox.Size = New System.Drawing.Size(502, 147)
        Me.SearchResultsListBox.TabIndex = 13
        '
        'm_SearchingForLabel
        '
        Me.m_SearchingForLabel.AutoSize = True
        Me.m_SearchingForLabel.Location = New System.Drawing.Point(12, 9)
        Me.m_SearchingForLabel.Name = "m_SearchingForLabel"
        Me.m_SearchingForLabel.Size = New System.Drawing.Size(76, 13)
        Me.m_SearchingForLabel.TabIndex = 18
        Me.m_SearchingForLabel.Text = "Searching for: "
        '
        'BtnDone
        '
        Me.BtnDone.Location = New System.Drawing.Point(439, 219)
        Me.BtnDone.Name = "BtnDone"
        Me.BtnDone.Size = New System.Drawing.Size(75, 23)
        Me.BtnDone.TabIndex = 19
        Me.BtnDone.Text = "Done!"
        Me.BtnDone.UseVisualStyleBackColor = True
        '
        'BtnNoFile
        '
        Me.BtnNoFile.Location = New System.Drawing.Point(15, 219)
        Me.BtnNoFile.Name = "BtnNoFile"
        Me.BtnNoFile.Size = New System.Drawing.Size(101, 23)
        Me.BtnNoFile.TabIndex = 20
        Me.BtnNoFile.Text = "No Matching File!"
        Me.BtnNoFile.UseVisualStyleBackColor = True
        '
        'm_itemsCountLabel
        '
        Me.m_itemsCountLabel.AutoSize = True
        Me.m_itemsCountLabel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.m_itemsCountLabel.Location = New System.Drawing.Point(12, 192)
        Me.m_itemsCountLabel.Name = "m_itemsCountLabel"
        Me.m_itemsCountLabel.Size = New System.Drawing.Size(43, 13)
        Me.m_itemsCountLabel.TabIndex = 21
        Me.m_itemsCountLabel.Text = "0 Items"
        '
        'FileSelectionForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(526, 254)
        Me.Controls.Add(Me.m_itemsCountLabel)
        Me.Controls.Add(Me.BtnNoFile)
        Me.Controls.Add(Me.BtnDone)
        Me.Controls.Add(Me.m_SearchingForLabel)
        Me.Controls.Add(Me.SearchResultsListBox)
        Me.Name = "FileSelectionForm"
        Me.Text = "FileSelectionForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents SearchResultsListBox As System.Windows.Forms.ListBox
    Public WithEvents m_SearchingForLabel As System.Windows.Forms.Label
    Private WithEvents BtnDone As System.Windows.Forms.Button
    Private WithEvents BtnNoFile As System.Windows.Forms.Button
    Public WithEvents m_itemsCountLabel As System.Windows.Forms.Label
End Class
