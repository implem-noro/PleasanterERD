[object]$connectionStringOrg = New-Object -TypeName System.Data.SqlClient.SqlConnectionStringBuilder
[object]$connectionStringOrg['Data Source'] = "(local)"
[object]$connectionStringOrg['Initial Catalog'] = "Implem.Pleasanter"
[object]$connectionStringOrg['UID'] = "sa"
[object]$connectionStringOrg['PWD'] = "password"
[object]$connectionStringDst = New-Object -TypeName System.Data.SqlClient.SqlConnectionStringBuilder
[object]$connectionStringDst['Data Source'] = "(local)"
[object]$connectionStringDst['Initial Catalog'] = "Implem.PleasanterERD"
[object]$connectionStringDst['UID'] = "sa"
[object]$connectionStringDst['PWD'] = "password"
[object]$columnNames = @{
    "IssueId" = "ID"
    "ResultId" = "ID"
    "Ver" = "バージョン"
    "Title" = "タイトル"
    "Body" = "内容"
    "StartTime" = "開始"
    "CompletionTime" = "完了"
    "WorkValue" = "作業量"
    "ProgressRate" = "進捗率"
    "RemainingWorkValue" = "残作業量"
    "Status" = "状況"
    "Manager" = "管理者"
    "Owner" = "担当者"
    "Comments" = "コメント"
    "CreatedTime" = "作成日時"
    "UpdatedTime" = "更新日時"
    "Creator" = "作成者"
    "Updator" = "更新者"
}
[object]$defaultIssuesColumn = @(
    "IssueId",
    "Ver",
    "Title",
    "Body",
    "StartTime",
    "CompletionTime",
    "WorkValue",
    "ProgressRate",
    "RemainingWorkValue",
    "Status",
    "Manager",
    "Owner",
    "Comments"
)
[object]$defaultResultsColumn = @(
    "ResultId",
    "Ver",
    "Title",
    "Body",
    "Status",
    "Manager",
    "Owner",
    "Comments"
)

function GetSite()
{
    [string]$sqlQuery = "select * from [Sites] where [ReferenceType] in ('Issues', 'Results');"
    [object]$resultsDataTable = New-Object System.Data.DataTable
    [object]$sqlConnection = New-Object System.Data.SQLClient.SQLConnection($connectionStringOrg)
    [object]$sqlCommand = New-Object System.Data.SQLClient.SQLCommand($sqlQuery, $sqlConnection)
    [object]$sqlConnection.Open()
    [object]$resultsDataTable.Load($sqlCommand.ExecuteReader())
    [object]$sqlConnection.Close()
    return $resultsDataTable
}

function MainteTable($sqlOperation, $row)
{
    [string]$referenceType = "$([string]$row.ReferenceType)"
    [object]$ss = ConvertFrom-Json ([string]$row.SiteSettings)
    [object]$sqlColumns = New-Object System.Collections.Generic.List[string]
    [object]$editorColumnsOrg = $ss.EditorColumnHash
    [object]$editorColumnsCnv = $ss.EditorColumnHash
    if ($editorColumnsOrg.Count -eq 0)
    {
        if ($referenceType -eq 'Issues')
        {
            $editorColumnsOrg = $defaultIssuesColumn
            $editorColumnsCnv = $defaultIssuesColumn
        }
        elseif ($referenceType -eq 'Results')
        {
            $editorColumnsOrg = $defaultResultsColumn
            $editorColumnsCnv = $defaultResultsColumn
        }
    }
    else
    {
        $editorColumnsOrg = New-Object System.Collections.ArrayList
        $editorColumnsCnv = New-Object System.Collections.ArrayList
        $editorColumnHash = (ConvertConvertFrom-JsonPSCustomObjectToHash $ss.EditorColumnHash)
        foreach ($key in $editorColumnHash.Keys)
        {
            foreach ($value in $editorColumnHash[$key])
            {
                [int]$ret = $editorColumnsOrg.Add($value)
            }
        }
        ConvertEditorColumns
    }
    foreach ($columnName in $editorColumnsCnv)
    {
        [string]$siteId = "$([string]$row.SiteId)"
        [string]$tableNameId = $siteId
        [string]$tableNameJpn = "$([string]$row.Title)" + "（" + $referenceType + "_" + $siteId + "）"
        switch ($sqlOperation)
        {
            'dropForeignKey' {DropForeignKey($row)}
            'createTable' {SetAlias($row)}
            'addForeignKey' {AddForeignKey($row)}
            'renameTableColumn' {SetAlias($row)}
        }
    }
    OperateForEachTable($row)
}

function ConvertEditorColumns($row)
{
    if ($referenceType -eq 'Issues')
    {
        if ($editorColumnsOrg -notcontains "IssueId")
        {
            [int]$ret = $editorColumnsCnv.Add("IssueId")
        }
    }
    elseif ($referenceType -eq 'Results')
    {
        if ($editorColumnsOrg -notcontains "ResultId")
        {
            [int]$ret = $editorColumnsCnv.Add("ResultId")
        }
    }
    foreach($editorColumn in $editorColumnsOrg) {
        if ( -not (
            ($editorColumn -match '_Section-[\d]') -or
            ($editorColumn -match '_Links-[\d]')))
        {
            [int]$ret = $editorColumnsCnv.Add($editorColumn)
        }
    }
}

function DropForeignKey($row)
{
    [string]$linkFromColumn = ""
    [string]$linkToSiteId = ""
    [string]$linkToSystemTable = ""
    foreach ($column in $ss.Links)
    {
        if ($column.ColumnName -eq $columnName)
        {
            $linkFromColumn = $column.ColumnName
            $linkToSiteId = $column.SiteId
            $linkToSystemTable = $column.TableName
            SqlDropForeignKey $row
        }
    }
}

function AddForeignKey($row)
{
    [string]$linkFromColumn = ""
    [string]$linkToSiteId = ""
    [string]$linkToSystemTable = ""
    foreach ($column in $ss.Links)
    {
        if ($column.ColumnName -eq $columnName)
        {
            $linkFromColumn = $column.ColumnName
            $linkToSiteId = $column.SiteId
            $linkToSystemTable = $column.TableName
            SqlChangeDataType $row
            SqlAddForeignKey $row
        }
    }
}

function SetAlias($row)
{
    [string]$alias = ""
    foreach ($column in $ss.Columns)
    {
        if ($column.ColumnName -eq $columnName)
        {
            $alias = $column.LabelText
        }
    }
    if ($alias -eq "")
    {
        $alias = $columnNames[$columnName]
    }
    if ($alias -eq "")
    {
        $alias = $columnName
    }
    switch ($sqlOperation)
    {
        'createTable' {CreateTable($row)}
        'renameTableColumn' {RenameTableColumn($row)}
    }
}

function OperateForEachTable($row)
{
    if ($sqlOperation -eq 'deleteTable')
    {
        SqlDeleteTable $row
    }
    elseif ($sqlOperation -eq 'createTable')
    {
        [string]$joinedSqlColumns = $sqlColumns -join ","
        SqlCreateTable $row $joinedSqlColumns
    }
    elseif ($sqlOperation -eq 'renameTableName')
    {
        [string]$oldTableName = $siteId
        [string]$newTableName = "$([string]$row.Title)" + "（" + $referenceType + "_" + $siteId + "）"
        SqlRenameTableName $row
    }
}

function CreateTable($row)
{
    $columnDefinition = GetColumnDefinition $columnName
    if (($columnName -eq "IssueId") -or ($columnName -eq "ResultId"))
    {
        $sqlColumns.Add("[Id]" + $columnDefinition)
    }
    else
    {
        $sqlColumns.Add("[" + $columnName + "]" + $columnDefinition)
    }
}

function RenameTableColumn($row)
{
    [string]$oldColumnName = ""
    [string]$newColumnName = ""
    if (($columnName -eq "IssueId") -or ($columnName -eq "ResultId"))
    {
        $oldColumnName = "Id"
        $newColumnName = $alias + "（" + $columnName + "）"
    }
    else
    {
        $oldColumnName = $columnName
        $newColumnName = $alias + "（" + $columnName + "）"
    }
    SqlRenameTableColumn $row
}

function GetColumnDefinition($columnName)
{
    switch -Regex ($columnName)
    {
        'SiteId' {'[bigint] NOT NULL'}
        'UpdatedTime' {'[datetime] NOT NULL'}
        'IssueId' {'[int] NOT NULL PRIMARY KEY'}
        'ResultId' {'[int] NOT NULL PRIMARY KEY'}
        'Ver' {'[int] NOT NULL'}
        'Title' {'[nvarchar](1024) NOT NULL'}
        'Body' {'[nvarchar](max) NULL'}
        'StartTime' {'[datetime] NULL'}
        'CompletionTime' {'[datetime] NOT NULL'}
        'WorkValue' {'[decimal](19, 4) NULL'}
        'ProgressRate' {'[decimal](4, 1) NULL'}
        'Status' {'[int] NOT NULL'}
        'Manager' {'[int] NULL'}
        'Owner' {'[int] NULL'}
        'Locked' {'[bit] NULL'}
        'Class[A-Z]'{'[nvarchar](1024) NULL'}
        'Class[\d{3}]'{'[nvarchar](1024) NULL'}
        'Num[A-Z]' {'[decimal](19, 4) NULL'}
        'Num[\d{3}]' {'[decimal](19, 4) NULL'}
        'Date[A-Z]' {'[datetime] NULL'}
        'Date[\d{3}]' {'[datetime] NULL'}
        'Description[A-Z]' {'[nvarchar](max) NULL'}
        'Description[\d{3}]' {'[nvarchar](max) NULL'}
        'Check[A-Z]' {'[bit] NULL'}
        'Check[\d{3}]' {'[bit] NULL'}
        'Attachments[A-Z]' {'[nvarchar](max) NULL'}
        'Attachments[\d{3}]' {'[nvarchar](max) NULL'}
        'Comments' {'[nvarchar](max) NULL'}
        'Creator' {'[int] NOT NULL'}
        'Updator' {'[int] NOT NULL'}
        'CreatedTime' {'[datetime] NOT NULL'}
    }
}

function SqlDeleteTable($row)
{
    "テーブル削除：" + $tableNameId + "," + $tableNameJpn
    [string]$sqlQuery = @"
        if exists (select 1 from sysobjects where id = object_id('$tableNameId'))
        drop table [$tableNameId];
        if exists (select 1 from sysobjects where id = object_id('$tableNameJpn'))
        drop table [$tableNameJpn];
"@
    RunSql($sqlQuery)
}

function SqlCreateTable($row, $joinedSqlColumns)
{
    "テーブル作成：" + $tableNameId + "," + $tableNameJpn
    [string]$sqlQuery = @"
        create table [$tableNameId] `($joinedSqlColumns`);
"@
    RunSql($sqlQuery)
}

function SqlDropForeignKey($row)
{
    [string]$foreignKeyName = GetForeignKeyName
    "外部キーの削除：" + $foreignKeyName
    [string]$sqlQuery = @"
        if exists (select 1 from sysobjects where id = object_id('$tableNameId'))
        if exists (select 1 from information_schema.referential_constraints
            where constraint_name = N'$foreignKeyName')
        alter table [$tableNameId] drop constraint $foreignKeyName;
        if exists (select 1 from sysobjects where id = object_id('$tableNameJpn'))
        if exists (select 1 from information_schema.referential_constraints
            where constraint_name = N'$foreignKeyName')
        alter table [$tableNameJpn] drop constraint $foreignKeyName;
"@
    RunSql($sqlQuery)
}

function SqlChangeDataType($row)
{
    "データ型の変更：" + $linkFromColumn
    [string]$sqlQuery = @"
        alter table [$tableNameId] alter column $linkFromColumn int;
"@
    RunSql($sqlQuery)
}

function SqlAddForeignKey($row)
{
    [string]$foreignKeyName = GetForeignKeyName
    [string]$foreignKeyLinkToTable = GetForeignKeyLinkToTable
    [string]$foreignKeyLinkToObject = GetForeignKeyLinkToObject
    "外部キーの追加：" + $foreignKeyName
    [string]$sqlQuery = @"
        if exists (select 1 from sysobjects where id = object_id('$foreignKeyLinkToTable'))
        alter table [$tableNameId] add constraint $foreignKeyName
            foreign key `($linkFromColumn`)
            references $foreignKeyLinkToObject;
"@
    RunSql($sqlQuery)
}

function SqlRenameTableColumn($row, $joinedSqlColumns)
{
    "テーブル列名の変更：" + $tableNameId + "," + $oldColumnName + " -> " + $newColumnName
    [string]$sqlQuery = @"
        EXEC sp_rename '$tableNameId.$oldColumnName', '$newColumnName', 'COLUMN';
"@
    RunSql($sqlQuery)
}

function SqlRenameTableName($row, $joinedSqlColumns)
{
    "テーブル名の変更：" + $oldTableName + " -> " + $newTableName
    [string]$sqlQuery = @"
        EXEC sp_rename '$oldTableName', '$newTableName', 'OBJECT';
"@
    RunSql($sqlQuery)
}

function RunSql($sqlQuery)
{
    [object]$sqlConnection = New-Object System.Data.SQLClient.SQLConnection($connectionStringDst)
    [object]$sqlCommand = New-Object System.Data.SQLClient.SQLCommand($sqlQuery, $sqlConnection)
    [object]$sqlConnection.Open()
    [int]$ret = $sqlCommand.ExecuteNonQuery()
    [object]$sqlConnection.Close()
}

function GetForeignKeyName($row)
{
    switch ($linkToSystemTable)
    {
        'Depts' {return "FK_" + $tableNameId + "_" + $linkFromColumn + "_Depts"}
        'Groups' {return "FK_" + $tableNameId + "_" + $linkFromColumn + "_Groups"}
        'Users' {return "FK_" + $tableNameId + "_" + $linkFromColumn + "_Users"}
        default {return "FK_" + $tableNameId + "_" + $linkFromColumn + "_" + $linkToSiteId}
    }
}

function GetForeignKeyLinkToTable($row)
{
    switch ($linkToSystemTable)
    {
        'Depts' {return "Depts"}
        'Groups' {return "Groups"}
        'Users' {return "Users"}
        default {return "$linkToSiteId"}
    }
}

function GetForeignKeyLinkToObject($row)
{
    switch ($linkToSystemTable)
    {
        'Depts' {return "[Depts] `(DeptId`)"}
        'Groups' {return "[Groups] `(GroupId`)"}
        'Users' {return "[Users] `(UserId`)"}
        default {return "[$linkToSiteId] `(Id`)"}
    }
}

function ConvertConvertFrom-JsonPSCustomObjectToHash($obj)
{
    $hash = @{}
    $obj | Get-Member -MemberType Properties | SELECT -exp "Name" | % {
        $hash[$_] = ($obj | SELECT -exp $_)
    }
    $hash
}

function Pause() {
    Write-Host "続行するには何かキーを押してください..." -NoNewLine
    [Console]::ReadKey() | Out-Null
}

#メイン処理
[object]$table = GetSite
#外部キーの削除
foreach ($row in $table) { MainteTable 'dropForeignKey' $row }
#Pause
#テーブルの削除
foreach ($row in $table) { MainteTable 'deleteTable' $row }
#Pause
#テーブルの作成
foreach ($row in $table) { MainteTable 'createTable' $row }
#Pause
#外部キーの追加
foreach ($row in $table) { MainteTable 'addForeignKey' $row }
#Pause
#テーブル列名の変更
foreach ($row in $table) { MainteTable 'renameTableColumn' $row }
#Pause
#テーブル名の変更
foreach ($row in $table) { MainteTable 'renameTableName' $row }