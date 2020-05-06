$FormInstructorUtilization = New-Object -TypeName System.Windows.Forms.Form
[System.Windows.Forms.GroupBox]$GroupBoxClassesLoaded = $null
[System.Windows.Forms.DataGridView]$DataGridViewClassesLoaded = $null
[System.Windows.Forms.Label]$LabelTotalEvents = $null
[System.Windows.Forms.Button]$ButtonChangeDataSrc = $null
[System.Windows.Forms.Button]$ButtonRemoveClassSched = $null
[System.Windows.Forms.Button]$ButtonImportSched = $null
[System.Windows.Forms.GroupBox]$GroupBoxFilters = $null
[System.Windows.Forms.Button]$ButtonInvert = $null
[System.Windows.Forms.Button]$ButtonSelDOD = $null
[System.Windows.Forms.Button]$ButtonFilteredGrid = $null
[System.Windows.Forms.DataGridView]$DataGridViewInstructors = $null
[System.Windows.Forms.Label]$LabelFilteredEvents = $null
[System.Windows.Forms.Label]$LabelClassFilter = $null
[System.Windows.Forms.ComboBox]$ComboBoxClassFilter = $null
[System.Windows.Forms.Label]$LabelCourseFilter = $null
[System.Windows.Forms.ComboBox]$ComboBoxCourseFilter = $null
[System.Windows.Forms.Label]$LabelEndFilter = $null
[System.Windows.Forms.Label]$LabelStartFilter = $null
[System.Windows.Forms.DateTimePicker]$DateTimePickerEndFilter = $null
[System.Windows.Forms.DateTimePicker]$DateTimePickerStartFilter = $null
[System.Windows.Forms.Label]$LabelInstructorFilter = $null
[System.Windows.Forms.GroupBox]$GroupBoxReports = $null
[System.Windows.Forms.Label]$LabelCRUtilRate = $null
[System.Windows.Forms.NumericUpDown]$NumericUpDownCRUtilRate = $null
[System.Windows.Forms.Label]$LabelInstAvail = $null
[System.Windows.Forms.NumericUpDown]$NumericUpDownInstAvail = $null
[System.Windows.Forms.Button]$ButtonMonthlyReport = $null
[System.Windows.Forms.Button]$ButtonQuarterlyReport = $null
[System.Windows.Forms.GroupBox]$GroupBoxSchedEvents = $null
[System.Windows.Forms.Button]$ButtoniCalSched = $null
[System.Windows.Forms.Button]$ButtonOutlookSched = $null
function InitializeComponent
{
$GroupBoxClassesLoaded = (New-Object -TypeName System.Windows.Forms.GroupBox)
$DataGridViewClassesLoaded = (New-Object -TypeName System.Windows.Forms.DataGridView)
$LabelTotalEvents = (New-Object -TypeName System.Windows.Forms.Label)
$ButtonChangeDataSrc = (New-Object -TypeName System.Windows.Forms.Button)
$ButtonRemoveClassSched = (New-Object -TypeName System.Windows.Forms.Button)
$ButtonImportSched = (New-Object -TypeName System.Windows.Forms.Button)
$GroupBoxFilters = (New-Object -TypeName System.Windows.Forms.GroupBox)
$ButtonSelDOD = (New-Object -TypeName System.Windows.Forms.Button)
$ButtonFilteredGrid = (New-Object -TypeName System.Windows.Forms.Button)
$DataGridViewInstructors = (New-Object -TypeName System.Windows.Forms.DataGridView)
$LabelFilteredEvents = (New-Object -TypeName System.Windows.Forms.Label)
$LabelClassFilter = (New-Object -TypeName System.Windows.Forms.Label)
$ComboBoxClassFilter = (New-Object -TypeName System.Windows.Forms.ComboBox)
$LabelCourseFilter = (New-Object -TypeName System.Windows.Forms.Label)
$ComboBoxCourseFilter = (New-Object -TypeName System.Windows.Forms.ComboBox)
$LabelEndFilter = (New-Object -TypeName System.Windows.Forms.Label)
$LabelStartFilter = (New-Object -TypeName System.Windows.Forms.Label)
$DateTimePickerEndFilter = (New-Object -TypeName System.Windows.Forms.DateTimePicker)
$DateTimePickerStartFilter = (New-Object -TypeName System.Windows.Forms.DateTimePicker)
$LabelInstructorFilter = (New-Object -TypeName System.Windows.Forms.Label)
$GroupBoxReports = (New-Object -TypeName System.Windows.Forms.GroupBox)
$LabelCRUtilRate = (New-Object -TypeName System.Windows.Forms.Label)
$NumericUpDownCRUtilRate = (New-Object -TypeName System.Windows.Forms.NumericUpDown)
$LabelInstAvail = (New-Object -TypeName System.Windows.Forms.Label)
$NumericUpDownInstAvail = (New-Object -TypeName System.Windows.Forms.NumericUpDown)
$ButtonMonthlyReport = (New-Object -TypeName System.Windows.Forms.Button)
$ButtonQuarterlyReport = (New-Object -TypeName System.Windows.Forms.Button)
$GroupBoxSchedEvents = (New-Object -TypeName System.Windows.Forms.GroupBox)
$ButtoniCalSched = (New-Object -TypeName System.Windows.Forms.Button)
$ButtonOutlookSched = (New-Object -TypeName System.Windows.Forms.Button)
$ButtonInvert = (New-Object -TypeName System.Windows.Forms.Button)
$GroupBoxClassesLoaded.SuspendLayout()
([System.ComponentModel.ISupportInitialize]$DataGridViewClassesLoaded).BeginInit()
$GroupBoxFilters.SuspendLayout()
([System.ComponentModel.ISupportInitialize]$DataGridViewInstructors).BeginInit()
$GroupBoxReports.SuspendLayout()
([System.ComponentModel.ISupportInitialize]$NumericUpDownCRUtilRate).BeginInit()
([System.ComponentModel.ISupportInitialize]$NumericUpDownInstAvail).BeginInit()
$GroupBoxSchedEvents.SuspendLayout()
$FormInstructorUtilization.SuspendLayout()
#
#GroupBoxClassesLoaded
#
$GroupBoxClassesLoaded.Controls.Add($DataGridViewClassesLoaded)
$GroupBoxClassesLoaded.Controls.Add($LabelTotalEvents)
$GroupBoxClassesLoaded.Controls.Add($ButtonChangeDataSrc)
$GroupBoxClassesLoaded.Controls.Add($ButtonRemoveClassSched)
$GroupBoxClassesLoaded.Controls.Add($ButtonImportSched)
$GroupBoxClassesLoaded.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$GroupBoxClassesLoaded.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]27,[System.Int32]31))
$GroupBoxClassesLoaded.Name = [System.String]'GroupBoxClassesLoaded'
$GroupBoxClassesLoaded.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]775,[System.Int32]256))
$GroupBoxClassesLoaded.TabIndex = [System.Int32]0
$GroupBoxClassesLoaded.TabStop = $false
$GroupBoxClassesLoaded.Text = [System.String]'Classes Loaded'
$GroupBoxClassesLoaded.UseCompatibleTextRendering = $true
#
#DataGridViewClassesLoaded
#
$DataGridViewClassesLoaded.AllowUserToAddRows = $false
$DataGridViewClassesLoaded.AllowUserToDeleteRows = $false
$DataGridViewClassesLoaded.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$DataGridViewClassesLoaded.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]204,[System.Int32]42))
$DataGridViewClassesLoaded.MultiSelect = $false
$DataGridViewClassesLoaded.Name = [System.String]'DataGridViewClassesLoaded'
$DataGridViewClassesLoaded.ReadOnly = $true
$DataGridViewClassesLoaded.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$DataGridViewClassesLoaded.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]387,[System.Int32]208))
$DataGridViewClassesLoaded.TabIndex = [System.Int32]5
#
#LabelTotalEvents
#
$LabelTotalEvents.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]292,[System.Int32]17))
$LabelTotalEvents.Name = [System.String]'LabelTotalEvents'
$LabelTotalEvents.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]147,[System.Int32]22))
$LabelTotalEvents.TabIndex = [System.Int32]4
$LabelTotalEvents.Text = [System.String]'Total Events: 0'
$LabelTotalEvents.TextAlign = [System.Drawing.ContentAlignment]::BottomRight
$LabelTotalEvents.UseCompatibleTextRendering = $true
#
#ButtonChangeDataSrc
#
$ButtonChangeDataSrc.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$ButtonChangeDataSrc.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]625,[System.Int32]42))
$ButtonChangeDataSrc.Name = [System.String]'ButtonChangeDataSrc'
$ButtonChangeDataSrc.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]128,[System.Int32]44))
$ButtonChangeDataSrc.TabIndex = [System.Int32]3
$ButtonChangeDataSrc.Text = [System.String]'Change Data Source Path'
$ButtonChangeDataSrc.UseCompatibleTextRendering = $true
$ButtonChangeDataSrc.UseVisualStyleBackColor = $true
$ButtonChangeDataSrc.add_Click($ChangeDataSource)
#
#ButtonRemoveClassSched
#
$ButtonRemoveClassSched.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$ButtonRemoveClassSched.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]24,[System.Int32]92))
$ButtonRemoveClassSched.Name = [System.String]'ButtonRemoveClassSched'
$ButtonRemoveClassSched.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]146,[System.Int32]44))
$ButtonRemoveClassSched.TabIndex = [System.Int32]2
$ButtonRemoveClassSched.Text = [System.String]'Delete Class Schedule'
$ButtonRemoveClassSched.UseCompatibleTextRendering = $true
$ButtonRemoveClassSched.UseVisualStyleBackColor = $true
$ButtonRemoveClassSched.add_Click($DeleteSched)
#
#ButtonImportSched
#
$ButtonImportSched.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$ButtonImportSched.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]24,[System.Int32]42))
$ButtonImportSched.Name = [System.String]'ButtonImportSched'
$ButtonImportSched.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]146,[System.Int32]44))
$ButtonImportSched.TabIndex = [System.Int32]1
$ButtonImportSched.Text = [System.String]'New / Update Class Schedule'
$ButtonImportSched.UseCompatibleTextRendering = $true
$ButtonImportSched.UseVisualStyleBackColor = $true
$ButtonImportSched.add_Click($ImportSched)
#
#GroupBoxFilters
#
$GroupBoxFilters.Controls.Add($ButtonInvert)
$GroupBoxFilters.Controls.Add($ButtonSelDOD)
$GroupBoxFilters.Controls.Add($ButtonFilteredGrid)
$GroupBoxFilters.Controls.Add($DataGridViewInstructors)
$GroupBoxFilters.Controls.Add($LabelFilteredEvents)
$GroupBoxFilters.Controls.Add($LabelClassFilter)
$GroupBoxFilters.Controls.Add($ComboBoxClassFilter)
$GroupBoxFilters.Controls.Add($LabelCourseFilter)
$GroupBoxFilters.Controls.Add($ComboBoxCourseFilter)
$GroupBoxFilters.Controls.Add($LabelEndFilter)
$GroupBoxFilters.Controls.Add($LabelStartFilter)
$GroupBoxFilters.Controls.Add($DateTimePickerEndFilter)
$GroupBoxFilters.Controls.Add($DateTimePickerStartFilter)
$GroupBoxFilters.Controls.Add($LabelInstructorFilter)
$GroupBoxFilters.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$GroupBoxFilters.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]27,[System.Int32]293))
$GroupBoxFilters.Name = [System.String]'GroupBoxFilters'
$GroupBoxFilters.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]775,[System.Int32]243))
$GroupBoxFilters.TabIndex = [System.Int32]1
$GroupBoxFilters.TabStop = $false
$GroupBoxFilters.Text = [System.String]'Filters'
$GroupBoxFilters.UseCompatibleTextRendering = $true
#
#ButtonSelDOD
#
$ButtonSelDOD.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]6,[System.Int32]57))
$ButtonSelDOD.Name = [System.String]'ButtonSelDOD'
$ButtonSelDOD.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]112,[System.Int32]23))
$ButtonSelDOD.TabIndex = [System.Int32]15
$ButtonSelDOD.Text = [System.String]'Select DOD Instr'
$ButtonSelDOD.UseCompatibleTextRendering = $true
$ButtonSelDOD.UseVisualStyleBackColor = $true
$ButtonSelDOD.add_Click($SelDOD)
#
#ButtonFilteredGrid
#
$ButtonFilteredGrid.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]573,[System.Int32]191))
$ButtonFilteredGrid.Name = [System.String]'ButtonFilteredGrid'
$ButtonFilteredGrid.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]109,[System.Int32]37))
$ButtonFilteredGrid.TabIndex = [System.Int32]14
$ButtonFilteredGrid.Text = [System.String]'View Events'
$ButtonFilteredGrid.UseCompatibleTextRendering = $true
$ButtonFilteredGrid.UseVisualStyleBackColor = $true
$ButtonFilteredGrid.add_Click($ViewEventGrid)
#
#DataGridViewInstructors
#
$DataGridViewInstructors.AllowUserToAddRows = $false
$DataGridViewInstructors.AllowUserToDeleteRows = $false
$DataGridViewInstructors.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
$DataGridViewInstructors.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$DataGridViewInstructors.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]124,[System.Int32]33))
$DataGridViewInstructors.Name = [System.String]'DataGridViewInstructors'
$DataGridViewInstructors.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$DataGridViewInstructors.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]348,[System.Int32]204))
$DataGridViewInstructors.TabIndex = [System.Int32]13
#
#LabelFilteredEvents
#
$LabelFilteredEvents.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]239,[System.Int32]7))
$LabelFilteredEvents.Name = [System.String]'LabelFilteredEvents'
$LabelFilteredEvents.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]233,[System.Int32]23))
$LabelFilteredEvents.TabIndex = [System.Int32]12
$LabelFilteredEvents.Text = [System.String]'Filtered Events: 0'
$LabelFilteredEvents.TextAlign = [System.Drawing.ContentAlignment]::BottomRight
$LabelFilteredEvents.UseCompatibleTextRendering = $true
$LabelFilteredEvents.Visible = $false
#
#LabelClassFilter
#
$LabelClassFilter.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]491,[System.Int32]152))
$LabelClassFilter.Name = [System.String]'LabelClassFilter'
$LabelClassFilter.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]64,[System.Int32]21))
$LabelClassFilter.TabIndex = [System.Int32]11
$LabelClassFilter.Text = [System.String]'Class'
$LabelClassFilter.TextAlign = [System.Drawing.ContentAlignment]::BottomRight
$LabelClassFilter.UseCompatibleTextRendering = $true
#
#ComboBoxClassFilter
#
$ComboBoxClassFilter.FormattingEnabled = $true
$ComboBoxClassFilter.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]561,[System.Int32]152))
$ComboBoxClassFilter.Name = [System.String]'ComboBoxClassFilter'
$ComboBoxClassFilter.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]121,[System.Int32]21))
$ComboBoxClassFilter.TabIndex = [System.Int32]10
#
#LabelCourseFilter
#
$LabelCourseFilter.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]491,[System.Int32]112))
$LabelCourseFilter.Name = [System.String]'LabelCourseFilter'
$LabelCourseFilter.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]64,[System.Int32]21))
$LabelCourseFilter.TabIndex = [System.Int32]9
$LabelCourseFilter.Text = [System.String]'Course'
$LabelCourseFilter.TextAlign = [System.Drawing.ContentAlignment]::BottomRight
$LabelCourseFilter.UseCompatibleTextRendering = $true
#
#ComboBoxCourseFilter
#
$ComboBoxCourseFilter.FormattingEnabled = $true
$ComboBoxCourseFilter.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]561,[System.Int32]112))
$ComboBoxCourseFilter.Name = [System.String]'ComboBoxCourseFilter'
$ComboBoxCourseFilter.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]121,[System.Int32]21))
$ComboBoxCourseFilter.TabIndex = [System.Int32]8
#
#LabelEndFilter
#
$LabelEndFilter.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$LabelEndFilter.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]489,[System.Int32]72))
$LabelEndFilter.Name = [System.String]'LabelEndFilter'
$LabelEndFilter.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]64,[System.Int32]21))
$LabelEndFilter.TabIndex = [System.Int32]6
$LabelEndFilter.Text = [System.String]'End'
$LabelEndFilter.TextAlign = [System.Drawing.ContentAlignment]::BottomRight
$LabelEndFilter.UseCompatibleTextRendering = $true
#
#LabelStartFilter
#
$LabelStartFilter.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$LabelStartFilter.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]489,[System.Int32]33))
$LabelStartFilter.Name = [System.String]'LabelStartFilter'
$LabelStartFilter.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]64,[System.Int32]21))
$LabelStartFilter.TabIndex = [System.Int32]5
$LabelStartFilter.Text = [System.String]'Start'
$LabelStartFilter.TextAlign = [System.Drawing.ContentAlignment]::BottomRight
$LabelStartFilter.UseCompatibleTextRendering = $true
#
#DateTimePickerEndFilter
#
$DateTimePickerEndFilter.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$DateTimePickerEndFilter.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]561,[System.Int32]72))
$DateTimePickerEndFilter.Name = [System.String]'DateTimePickerEndFilter'
$DateTimePickerEndFilter.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]200,[System.Int32]21))
$DateTimePickerEndFilter.TabIndex = [System.Int32]3
$DateTimePickerEndFilter.Value = (New-Object -TypeName System.DateTime -ArgumentList @([System.Int32]2020,[System.Int32]4,[System.Int32]20,[System.Int32]0,[System.Int32]0,[System.Int32]0,[System.Int32]0))
#
#DateTimePickerStartFilter
#
$DateTimePickerStartFilter.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$DateTimePickerStartFilter.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]561,[System.Int32]33))
$DateTimePickerStartFilter.Name = [System.String]'DateTimePickerStartFilter'
$DateTimePickerStartFilter.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]200,[System.Int32]21))
$DateTimePickerStartFilter.TabIndex = [System.Int32]2
$DateTimePickerStartFilter.Value = (New-Object -TypeName System.DateTime -ArgumentList @([System.Int32]2020,[System.Int32]4,[System.Int32]20,[System.Int32]0,[System.Int32]0,[System.Int32]0,[System.Int32]0))
#
#LabelInstructorFilter
#
$LabelInstructorFilter.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$LabelInstructorFilter.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]18,[System.Int32]33))
$LabelInstructorFilter.Name = [System.String]'LabelInstructorFilter'
$LabelInstructorFilter.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]100,[System.Int32]21))
$LabelInstructorFilter.TabIndex = [System.Int32]1
$LabelInstructorFilter.Text = [System.String]'Instructors'
$LabelInstructorFilter.TextAlign = [System.Drawing.ContentAlignment]::BottomRight
$LabelInstructorFilter.UseCompatibleTextRendering = $true
#
#GroupBoxReports
#
$GroupBoxReports.Controls.Add($LabelCRUtilRate)
$GroupBoxReports.Controls.Add($NumericUpDownCRUtilRate)
$GroupBoxReports.Controls.Add($LabelInstAvail)
$GroupBoxReports.Controls.Add($NumericUpDownInstAvail)
$GroupBoxReports.Controls.Add($ButtonMonthlyReport)
$GroupBoxReports.Controls.Add($ButtonQuarterlyReport)
$GroupBoxReports.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$GroupBoxReports.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]27,[System.Int32]542))
$GroupBoxReports.Name = [System.String]'GroupBoxReports'
$GroupBoxReports.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]402,[System.Int32]116))
$GroupBoxReports.TabIndex = [System.Int32]2
$GroupBoxReports.TabStop = $false
$GroupBoxReports.Text = [System.String]'Reports'
$GroupBoxReports.UseCompatibleTextRendering = $true
#
#LabelCRUtilRate
#
$LabelCRUtilRate.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]0,[System.Int32]18))
$LabelCRUtilRate.Name = [System.String]'LabelCRUtilRate'
$LabelCRUtilRate.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]112,[System.Int32]23))
$LabelCRUtilRate.TabIndex = [System.Int32]5
$LabelCRUtilRate.Text = [System.String]'CR Util Rate %'
$LabelCRUtilRate.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$LabelCRUtilRate.UseCompatibleTextRendering = $true
#
#NumericUpDownCRUtilRate
#
$NumericUpDownCRUtilRate.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]124,[System.Int32]20))
$NumericUpDownCRUtilRate.Name = [System.String]'NumericUpDownCRUtilRate'
$NumericUpDownCRUtilRate.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]57,[System.Int32]21))
$NumericUpDownCRUtilRate.TabIndex = [System.Int32]4
$NumericUpDownCRUtilRate.Value = (New-Object -TypeName System.Decimal -ArgumentList @(,[System.Int32[]]@([System.Int32]37,[System.Int32]0,[System.Int32]0,[System.Int32]0)))
#
#LabelInstAvail
#
$LabelInstAvail.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]174,[System.Int32]20))
$LabelInstAvail.Name = [System.String]'LabelInstAvail'
$LabelInstAvail.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]153,[System.Int32]23))
$LabelInstAvail.TabIndex = [System.Int32]3
$LabelInstAvail.Text = [System.String]'Instructors Available'
$LabelInstAvail.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$LabelInstAvail.UseCompatibleTextRendering = $true
#
#NumericUpDownInstAvail
#
$NumericUpDownInstAvail.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]333,[System.Int32]20))
$NumericUpDownInstAvail.Minimum = (New-Object -TypeName System.Decimal -ArgumentList @(,[System.Int32[]]@([System.Int32]2,[System.Int32]0,[System.Int32]0,[System.Int32]0)))
$NumericUpDownInstAvail.Name = [System.String]'NumericUpDownInstAvail'
$NumericUpDownInstAvail.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]63,[System.Int32]21))
$NumericUpDownInstAvail.TabIndex = [System.Int32]2
$NumericUpDownInstAvail.UpDownAlign = [System.Windows.Forms.LeftRightAlignment]::Left
$NumericUpDownInstAvail.Value = (New-Object -TypeName System.Decimal -ArgumentList @(,[System.Int32[]]@([System.Int32]22,[System.Int32]0,[System.Int32]0,[System.Int32]0)))
#
#ButtonMonthlyReport
#
$ButtonMonthlyReport.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]218,[System.Int32]53))
$ButtonMonthlyReport.Name = [System.String]'ButtonMonthlyReport'
$ButtonMonthlyReport.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]142,[System.Int32]56))
$ButtonMonthlyReport.TabIndex = [System.Int32]1
$ButtonMonthlyReport.Text = [System.String]'Monthly Rollup'
$ButtonMonthlyReport.UseCompatibleTextRendering = $true
$ButtonMonthlyReport.UseVisualStyleBackColor = $true
$ButtonMonthlyReport.add_Click($MonthlyReport)
#
#ButtonQuarterlyReport
#
$ButtonQuarterlyReport.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]35,[System.Int32]53))
$ButtonQuarterlyReport.Name = [System.String]'ButtonQuarterlyReport'
$ButtonQuarterlyReport.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]146,[System.Int32]56))
$ButtonQuarterlyReport.TabIndex = [System.Int32]0
$ButtonQuarterlyReport.Text = [System.String]'Quarterly Rollup'
$ButtonQuarterlyReport.UseCompatibleTextRendering = $true
$ButtonQuarterlyReport.UseVisualStyleBackColor = $true
$ButtonQuarterlyReport.add_Click($QuarterlyReport)
#
#GroupBoxSchedEvents
#
$GroupBoxSchedEvents.Controls.Add($ButtoniCalSched)
$GroupBoxSchedEvents.Controls.Add($ButtonOutlookSched)
$GroupBoxSchedEvents.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]8.25,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$GroupBoxSchedEvents.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]435,[System.Int32]542))
$GroupBoxSchedEvents.Name = [System.String]'GroupBoxSchedEvents'
$GroupBoxSchedEvents.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]367,[System.Int32]116))
$GroupBoxSchedEvents.TabIndex = [System.Int32]3
$GroupBoxSchedEvents.TabStop = $false
$GroupBoxSchedEvents.Text = [System.String]'Schedule Events'
$GroupBoxSchedEvents.UseCompatibleTextRendering = $true
#
#ButtoniCalSched
#
$ButtoniCalSched.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]208,[System.Int32]32))
$ButtoniCalSched.Name = [System.String]'ButtoniCalSched'
$ButtoniCalSched.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]137,[System.Int32]56))
$ButtoniCalSched.TabIndex = [System.Int32]1
$ButtoniCalSched.Text = [System.String]'Create iCal File'
$ButtoniCalSched.UseCompatibleTextRendering = $true
$ButtoniCalSched.UseVisualStyleBackColor = $true
$ButtoniCalSched.add_Click($CreateiCal)
#
#ButtonOutlookSched
#
$ButtonOutlookSched.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]23,[System.Int32]32))
$ButtonOutlookSched.Name = [System.String]'ButtonOutlookSched'
$ButtonOutlookSched.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]131,[System.Int32]56))
$ButtonOutlookSched.TabIndex = [System.Int32]0
$ButtonOutlookSched.Text = [System.String]'Create Outlook Schedule'
$ButtonOutlookSched.UseCompatibleTextRendering = $true
$ButtonOutlookSched.UseVisualStyleBackColor = $true
$ButtonOutlookSched.add_Click($OpenOutlookForm)
#
#ButtonInvert
#
$ButtonInvert.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]6,[System.Int32]86))
$ButtonInvert.Name = [System.String]'ButtonInvert'
$ButtonInvert.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]112,[System.Int32]23))
$ButtonInvert.TabIndex = [System.Int32]16
$ButtonInvert.Text = [System.String]'Invert Selection'
$ButtonInvert.UseCompatibleTextRendering = $true
$ButtonInvert.UseVisualStyleBackColor = $true
$ButtonInvert.add_Click($InvertSelection)
#
#FormInstructorUtilization
#
$FormInstructorUtilization.ClientSize = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]829,[System.Int32]671))
$FormInstructorUtilization.Controls.Add($GroupBoxSchedEvents)
$FormInstructorUtilization.Controls.Add($GroupBoxReports)
$FormInstructorUtilization.Controls.Add($GroupBoxFilters)
$FormInstructorUtilization.Controls.Add($GroupBoxClassesLoaded)
$FormInstructorUtilization.Font = (New-Object -TypeName System.Drawing.Font -ArgumentList @([System.String]'Tahoma',[System.Single]9.75,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Point,([System.Byte][System.Byte]0)))
$FormInstructorUtilization.Text = [System.String]'Instructor Utilization'
$FormInstructorUtilization.add_Shown($InstrUtilizationLoaded)
$FormInstructorUtilization.add_Click($QuarterlyReport)
$GroupBoxClassesLoaded.ResumeLayout($false)
([System.ComponentModel.ISupportInitialize]$DataGridViewClassesLoaded).EndInit()
$GroupBoxFilters.ResumeLayout($false)
([System.ComponentModel.ISupportInitialize]$DataGridViewInstructors).EndInit()
$GroupBoxReports.ResumeLayout($false)
([System.ComponentModel.ISupportInitialize]$NumericUpDownCRUtilRate).EndInit()
([System.ComponentModel.ISupportInitialize]$NumericUpDownInstAvail).EndInit()
$GroupBoxSchedEvents.ResumeLayout($false)
$FormInstructorUtilization.ResumeLayout($false)
Add-Member -InputObject $FormInstructorUtilization -Name base -Value $base -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name GroupBoxClassesLoaded -Value $GroupBoxClassesLoaded -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name DataGridViewClassesLoaded -Value $DataGridViewClassesLoaded -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name LabelTotalEvents -Value $LabelTotalEvents -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name ButtonChangeDataSrc -Value $ButtonChangeDataSrc -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name ButtonRemoveClassSched -Value $ButtonRemoveClassSched -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name ButtonImportSched -Value $ButtonImportSched -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name GroupBoxFilters -Value $GroupBoxFilters -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name ButtonInvert -Value $ButtonInvert -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name ButtonSelDOD -Value $ButtonSelDOD -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name ButtonFilteredGrid -Value $ButtonFilteredGrid -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name DataGridViewInstructors -Value $DataGridViewInstructors -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name LabelFilteredEvents -Value $LabelFilteredEvents -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name LabelClassFilter -Value $LabelClassFilter -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name ComboBoxClassFilter -Value $ComboBoxClassFilter -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name LabelCourseFilter -Value $LabelCourseFilter -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name ComboBoxCourseFilter -Value $ComboBoxCourseFilter -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name LabelEndFilter -Value $LabelEndFilter -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name LabelStartFilter -Value $LabelStartFilter -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name DateTimePickerEndFilter -Value $DateTimePickerEndFilter -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name DateTimePickerStartFilter -Value $DateTimePickerStartFilter -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name LabelInstructorFilter -Value $LabelInstructorFilter -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name GroupBoxReports -Value $GroupBoxReports -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name LabelCRUtilRate -Value $LabelCRUtilRate -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name NumericUpDownCRUtilRate -Value $NumericUpDownCRUtilRate -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name LabelInstAvail -Value $LabelInstAvail -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name NumericUpDownInstAvail -Value $NumericUpDownInstAvail -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name ButtonMonthlyReport -Value $ButtonMonthlyReport -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name ButtonQuarterlyReport -Value $ButtonQuarterlyReport -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name GroupBoxSchedEvents -Value $GroupBoxSchedEvents -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name ButtoniCalSched -Value $ButtoniCalSched -MemberType NoteProperty
Add-Member -InputObject $FormInstructorUtilization -Name ButtonOutlookSched -Value $ButtonOutlookSched -MemberType NoteProperty
}
. InitializeComponent
