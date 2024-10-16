###############################################################
###############################################################
function Assert-ModuleExists([string]$ModuleName) {
    $module = Get-Module $ModuleName -ListAvailable -ErrorAction SilentlyContinue
    if (!$module) {
        Write-Host "Installing module $ModuleName ..."
        Install-Module -Name $ModuleName -Force -Scope Allusers
        Write-Host "Module installed"
    }
    elseif ($module.Version -ne '1.0.0' -and $module.Version -le '1.0.410') {
        Write-Host "Updating module $ModuleName ..."
        Update-Module -Name $ModuleName -Force -ErrorAction Stop
        Write-Host "Module updated"
    }
}
## Install-Module -Name MicrosoftPowerBIMgmt
Assert-ModuleExists -ModuleName "MicrosoftPowerBIMgmt"
## Connect to the PowerBi Service
Connect-PowerBIServiceAccount

#########################Paremeters############################
## Options "bi_for_sccm" or "bi_for_intune or bi_for_defender"
#$powerbiapp = "bi_for_intune"
$AppArray = @()
$tempapp = New-Object -TypeName PSObject
$tempapp | Add-Member -MemberType NoteProperty -Name "App" -Value "bi_for_intune" -Force
$AppArray += $tempapp
$tempapp = New-Object -TypeName PSObject
$tempapp | Add-Member -MemberType NoteProperty -Name "App" -Value "bi_for_sccm" -Force
$AppArray += $tempapp
$tempapp = New-Object -TypeName PSObject
$tempapp | Add-Member -MemberType NoteProperty -Name "App" -Value "bi_for_defender" -Force
$AppArray += $tempapp
$AppSelect = $AppArray | Out-GridView -Title 'Choose an app to backup' -PassThru
$powerbiapp = $AppSelect.App

Write-Host "App to backup '" $powerbiapp "'"

###############################################################
#Find the Workspaces
$AllWorkspaces = Get-PowerBIWorkspace ### FILTER TO SHOW ONLY BI FOR INTUNE Workspaces REMOVE TO SHOW OTHERS ##### | Where-Object {$_.Name -like "*Intune*"}
## Select the source workspace
$SourceWorkspace = $AllWorkspaces | Out-GridView -Title 'Choose a source workspace' -PassThru
Write-Host "Source Workspace '" $SourceWorkspace.Name "'"

##Find the source dataset
$SourceDataset = Get-PowerBIDataset -WorkspaceId $SourceWorkspace.Id | Where-Object {$_.Name -eq $powerbiapp}

if($SourceDataset -ne $null){
    
    ## Select the destination workspace
    $DestinationWorkspace = $AllWorkspaces | Where-Object {$_.Id -ne $SourceWorkspace.Id} | Out-GridView -Title 'Choose a destination workspace' -PassThru
    Write-Host "Destination Workspace '" $DestinationWorkspace.Name "'"

    ##Find the destination dataset
    $DestinationDataset = Get-PowerBIDataset -WorkspaceId $DestinationWorkspace.Id | Where-Object {$_.Name -eq $powerbiapp}

    
    if($DestinationDataset -ne $null){


        ###################################
        ##Destination Workspace
        ###################################
        ## Find the custom reports in the source workspace
        $SourceReports = Get-PowerBIReport -WorkspaceId $SourceWorkspace.Id

        ## Copy the source reports to the destination workspace using the destination dataset
        $SourceReportsUniqueID = @()
        $SourceReportIDMapping = @{} 
        Foreach ($Report in $SourceReports){
            if(-Not($Report.Id -in $SourceReportsUniqueID)){
                $SourceReportsUniqueID += @($Report.Id)
                if($Report.Name -eq $powerbiapp){
                    $dt = Get-Date -Format "yyyy-MM-dd"
                    $NewReportName = $Report.Name + " " + $dt
                }else{
                    $NewReportName = $Report.Name
                }
                Write-host "About to copy Report '" $Report.Name "' to '" $NewReportName "' from Workspace '" $SourceWorkspace.Name "' to Workspace '" $DestinationWorkspace.Name "'"
                if($Report.DatasetId -eq $SourceDataset.Id){
                    $NewReport = Copy-PowerBIReport -Name $NewReportName -Id $Report.Id -WorkspaceId $SourceWorkspace.Id -TargetWorkspaceId $DestinationWorkspace.Id -TargetDatasetId $DestinationDataset.Id
                }else{
                    $NewReport = Copy-PowerBIReport -Name $NewReportName -Id $Report.Id -WorkspaceId $SourceWorkspace.Id -TargetWorkspaceId $DestinationWorkspace.Id -TargetDatasetId $Report.DatasetId
                }
                if ($NewReport){
                    $SourceReportIDMapping[$Report.Id] = [guid]$NewReport.Id
                }
            }
        } 

        ## Find the custom dashboards in the source workspace
        $SourceDashboards = Get-PowerBIDashboard -WorkspaceId $SourceWorkspace.Id

        ## Copy the source dashboards to the destination workspace using the destination reports
        $SourceDashboardsUniqueID = @()

        Foreach ($Dashboard in $SourceDashboards) {
            if(-Not($Dashboard.Id -in $SourceDashboardsUniqueID)){
                $SourceDashboardsUniqueID += @($Dashboard.Id)
                Write-Host "About to copy Dashboard '" $Dashboard.Name "' from Workspace '" $SourceWorkspace.Name "' to Workspace '" $DestinationWorkspace.Name "'"
                $NewDashboard = New-PowerBIDashboard -Name $Dashboard.Name -WorkspaceId $DestinationWorkspace.Id
                $SourceDashboardTiles = Get-PowerBITile -WorkspaceId $SourceWorkspace.Id -DashboardId $Dashboard.Id

                $TilesUniqueID = @()
                Foreach ($Tile in $SourceDashboardTiles) {
                    if(-Not($Tile.Id -in $TilesUniqueID)){
                        $TilesUniqueID += @($Tile.Id)
                        Write-Host "About to copy Tile '" $Tile.Id "' to Dashboard '" $NewDashboard.Name "in Workspace '" $DestinationWorkspace.Name "'"
                        if($SourceReportIDMapping[[guid]$Tile.ReportId]){
                            $NewTile = Copy-PowerBITile -WorkspaceId $SourceWorkspace.Id -DashboardId $Dashboard.Id -TileId $Tile.Id -TargetDashboardId $NewDashboard.Id -TargetWorkspaceId $DestinationWorkspace.Id -TargetReportId $SourceReportIDMapping[[guid]$Tile.ReportId] ##-TargetDatasetId $tile_target_dataset_id        
                        }else{
                            $NewTile = Copy-PowerBITile -WorkspaceId $SourceWorkspace.Id -DashboardId $Dashboard.Id -TileId $Tile.Id -TargetDashboardId $NewDashboard.Id -TargetWorkspaceId $DestinationWorkspace.Id -TargetReportId $Tile.ReportId ##-TargetDatasetId $tile_target_dataset_id                
                        }
                    }
                }
            }
        }

        ###################################
        ##Other Workspaces
        ###################################
        if($DestinationDataset.Id){
            ## Find the custom reports in the other workspaces
            $OtherWorkspaces = $AllWorkspaces | Where-Object {$_.Id -ne $SourceWorkspace.Id} | Where-Object {$_.Id -ne $DestinationWorkspace.Id}

            $OtherWorkspacesUniqueID = @()

            ## Copy the custom reports to the same workspace using the destination dataset
            Foreach ($OtherWorkspace in $OtherWorkspaces){
                if(-Not($OtherWorkspace.Id -in $OtherWorkspacesUniqueID)){
                    $OtherWorkspacesUniqueID += @($OtherWorkspace.Id)

                    $OtherReports = Get-PowerBIReport -WorkspaceId $OtherWorkspace.Id | Where-Object {$_.DatasetId -eq $SourceDataset.Id}

                    $OtherReportsUniqueID = @()
                    $OtherReportIDMapping = @{} 
                    Foreach ($OtherReport in $OtherReports){
                        if(-Not($OtherReport.Id -in $OtherReportsUniqueID)){
                            $OtherReportsUniqueID += @($OtherReport.Id)
                            $dt = Get-Date -Format "yyyy-MM-dd"
                            $BackupReportName = $OtherReport.Name + " " + $dt
                            Write-host "About to copy Report '" $OtherReport.Name "' to '" $BackupReportName "' in Workspace '" $OtherWorkspace.Name "'"
                            $NewOtherReport = Copy-PowerBIReport -Name $BackupReportName -Id $OtherReport.Id -WorkspaceId $OtherWorkspace.Id -TargetWorkspaceId $OtherWorkspace.Id -TargetDatasetId $DestinationDataset.Id
                            $OtherReportIDMapping[$OtherReport.Id] = [guid]$NewOtherReport.Id
                        }
                    }

                    ## Find the custom dashboards in the source workspace
                    $OtherDashboards = Get-PowerBIDashboard -WorkspaceId $OtherWorkspace.Id

                    ## Copy the custom dashboards to the same workspace using the copy reports
                    $OtherDashboardsUniqueID = @()

                    Foreach ($OtherDashboard in $OtherDashboards) {
                        if(-Not($OtherDashboard.Id -in $OtherDashboardsUniqueID)){
                            $OtherDashboardsUniqueID += @($OtherDashboard.Id)
                            $OtherDashboardTiles = Get-PowerBITile -WorkspaceId $OtherWorkspace.Id -DashboardId $OtherDashboard.Id
                
                            $CopyOtherDashboard = 0
                            $OtherTilesUniqueID = @()
                            Foreach ($OtherTile in $OtherDashboardTiles) {
                                if(-Not($OtherTile.Id -in $OtherTilesUniqueID)){
                                    $OtherTilesUniqueID += @($OtherTile.Id)
                                    if($OtherTile.DatasetId -eq $SourceDataset.Id){
                                        #This tile contains the Source Dataset
                                        $CopyOtherDashboard = 1
                                        break
                                    }
                                }
                            }
                            if($CopyOtherDashboard){
                                $dt = Get-Date -Format "yyyy-MM-dd"
                                $BackupDashboard = $OtherDashboard.Name + " " + $dt
                                Write-Host "About to copy Dashboard '" $OtherDashboard.Name "' to '" $BackupDashboard "' in Workspace '" $OtherWorkspace.Name "'"
                                $NewOtherDashboard = New-PowerBIDashboard -Name $BackupDashboard -WorkspaceId $OtherWorkspace.Id

                                $OtherTilesUniqueID = @()
                                Foreach ($OtherTile in $OtherDashboardTiles) {
                                    if(-Not($OtherTile.Id -in $OtherTilesUniqueID)){
                                        $OtherTilesUniqueID += @($OtherTile.Id)
                                        Write-Host "About to copy Tile '" $OtherTile.Id "' to Dashboard '" $NewOtherDashboard.Name "' in Workspace '" $OtherWorkspace.Name "'"
                                        if($OtherReportIDMapping[[guid]$OtherTile.ReportId]){
                                            $NewOtherTile = Copy-PowerBITile -WorkspaceId $OtherWorkspace.Id -DashboardId $OtherDashboard.Id -TileId $OtherTile.Id -TargetDashboardId $NewOtherDashboard.Id -TargetReportId $OtherReportIDMapping[[guid]$OtherTile.ReportId]
                                        }else{
                                            $NewOtherTile = Copy-PowerBITile -WorkspaceId $OtherWorkspace.Id -DashboardId $OtherDashboard.Id -TileId $OtherTile.Id -TargetDashboardId $NewOtherDashboard.Id -TargetReportId $OtherTile.ReportId
                                        }
                                    }
                                }
                            }               
                        }
                    }

                }
            }
        }

    }else{
        Write-Host "App '" $powerbiapp "' not found in destination workspace '" $DestinationDataset.Name "'"
    }

}else{
    Write-Host "App '" $powerbiapp "' not found in source workspace '" $SourceWorkspace.Name "'"
}
