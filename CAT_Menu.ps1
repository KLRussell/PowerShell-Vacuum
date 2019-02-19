# Created by Sam Escolas in 2016. Edited & Upgraded by Kevin Russell & Shane Johnson on 1/24/2019

#Global Variables (These cannot be changed)
$Global:SourceDir = (split-path $SCRIPT:MyInvocation.MyCommand.Path -parent);

$Global:SourceCodeDir = "$Global:SourceDir\03_Source_Code";
$Global:CATButtonsDir = "$Global:SourceCodeDir\01_CAT_Menu_Buttons";
$Global:AdminDir = "$Global:SourceDir\01_Updates\05_Admin\";

[int]$Global:MenuItem = 0;
[int]$Global:MenuStart = 0;
[int]$Global:MenuEnd = 0;

<#
    For functions 2-28, files will need to be re-located according to the filepath
#>
FUNCTION Run($show_menu=$true)
{
    If($show_menu)
    
    {Write-Host @"
    
Choose one of the following to run:
        
1`tVacuum Automation
2`t(Remote) Restart Vacuum
3`t(Remote) Stop Vacuum
4`t(Remote) Vacuum Status
5`tMasterCost OCCs
6`tGranite Status (Close Escalate)
7`tPaperCost
8`tDispute Rejections
9`tDispute Staging
10`tDispute Log
11`tDispute Log2
12`tDispute_Log_CG.xlsm
13`tDispute_Log_CT.xlsm
14`tDispute_Log_DM.xlsm
15`tDispute_Log_DW.xlsm
16`tDispute_Log_JB.xlsm
17`tDispute_Log_JL.xlsm
18`tDispute_Log_JQ.xlsm
19`tDispute_Log_KR.xlsm
20`tDispute Log_LPC.xlsm
21`tDispute_Log_LV.xlsm
22`tDispute_Log_MM.xlsm
23`tDispute_Log_ProductTeam.xlsm
24`tDispute_Log_PW.xlsm
25`tDispute_Log_RAT.xlsm
26`tDispute_Log_AD.xlsm
27`tDispute_Log_SL.xlsm
28`tDispute Log_TAX.xlsm
29`tDispute_Log_TS.xlsm
30`tDispute_Log_VT.xlsm
31`tDispute_Log_GT.xlsm
32`tMultiple
33`tCancel
99`tExit

        
"@

}

    if ($Global:MenuItem -eq 0)
    {
        $Global:MenuItem = Read-Host ">>> ";
    }

    Switch($Global:MenuItem)
    {
        {$_ -Eq "1" -or $_ -Match "vac"}
        {
            if (Vacuum-Online)
            {
                write-host "Error! Vacuum is currently running on $((cat -delimiter "," -path  (join-path $Global:AdminDir "VACUUM_ONLINE.TXT"))[0])'s machine";
            }
            else
            {
                write-host "Starting Vacuum....";
                Start-Process powershell.exe "& '$Global:SourceDir\03_Source_Code\Vacuum_Automation.ps1'" -WindowStyle Minimized;
            }
        }
        {$_ -Eq "2" -or $_ -Match "res"}
        {
            Remote-Command -Command "Restart" -Message "Sending remote command to current Vacuum user to restart Vacuum";
            Read-Host "Press Enter to exit..." | Out-Null;
        }
        {$_ -Eq "3" -or $_ -Match "stop"}
        {
            Remote-Command -Command "Shutdown" -Message "Sending remote command to current Vacuum user to stop Vacuum";
            Read-Host "Press Enter to exit..." | Out-Null;
        }
        {$_ -Eq "4" -or $_ -Match "stat"}
        {
            Remote-Command -Command "Status" -Message "Sending remote command to current Vacuum user to gather Vacuum status";
            Read-Host "Press Enter to exit..." | Out-Null;
        }
        {$_ -Eq "5" -or $_ -Match "occ"}
        {
            Start Powershell "$Global:CATButtonsDir\mastercost_occ.ps1; Read-Host 'success!`n`npress enter to exit'";
        }
        {$_ -Eq "6" -or $_ -Match "granite.*stat.*"}
        {
            Start Powershell "$Global:CATButtonsDir\granite_status.ps1; Read-Host 'success!`n`npress enter to exit'";
        }
        {$_ -Eq "7" -or $_ -Match "paper"}
        {
            Start Powershell "$Global:CATButtonsDir\paper_cost.ps1; Read-Host 'success!`n`npress enter to exit'"
        }
        {$_ -Eq "8" -or $_ -Match "reject"}
        {
            Start Powershell "$Global:CATButtonsDir\batch_rejections.ps1; Read-Host 'success!`n`npress enter to exit'";
        }
        {$_ -Eq "9" -or $_ -Match "dispute.*stag.*"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_staging.ps1; Read-Host 'success!`n`npress enter to exit'";
        }
        {$_ -Eq "10" -or $_ -Match "dispute.*log"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log.ps1; Read-Host 'success!`n`npress enter to exit'";
        }
        
        {$_ -Eq "11" -or $_ -Match "dispute.*log2"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log2.ps1; Read-Host 'success!`n`npress enter to exit'";
        }        
        
        {$_ -Eq "12" -or $_ -Match "dispute.*logCG"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_cg.ps1; Read-Host 'success!`n`npress enter to exit'";
        }        
        
        {$_ -Eq "13" -or $_ -Match "dispute.*logCT"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_ct.ps1; Read-Host 'success!`n`npress enter to exit'";
        }        

        {$_ -Eq "14" -or $_ -Match "dispute.*logDM"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_dm.ps1; Read-Host 'success!`n`npress enter to exit'";
        }

        {$_ -Eq "15" -or $_ -Match "dispute.*logDW"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_dw.ps1; Read-Host 'success!`n`npress enter to exit'";
        }

        {$_ -Eq "16" -or $_ -Match "dispute.*logJB"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_jb.ps1; Read-Host 'success!`n`npress enter to exit'";
        }

        {$_ -Eq "17" -or $_ -Match "dispute.*logJL"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_jl.ps1; Read-Host 'success!`n`npress enter to exit'";
        }

        {$_ -Eq "18" -or $_ -Match "dispute.*logJQ"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_jq.ps1; Read-Host 'success!`n`npress enter to exit'";
        }

        {$_ -Eq "19" -or $_ -Match "dispute.*logKR"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_kr.ps1; Read-Host 'success!`n`npress enter to exit'";
        }

        {$_ -Eq "20" -or $_ -Match "dispute.*logLPC"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_lpc.ps1; Read-Host 'success!`n`npress enter to exit'";
        }

        {$_ -Eq "21" -or $_ -Match "dispute.*logLV"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_lv.ps1; Read-Host 'success!`n`npress enter to exit'";
        }
       
        {$_ -Eq "22" -or $_ -Match "dispute.*logMM"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_mm.ps1; Read-Host 'success!`n`npress enter to exit'";
        }
       
        {$_ -Eq "23" -or $_ -Match "dispute.*logProductTeam"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_productteam.ps1; Read-Host 'success!`n`npress enter to exit'";
        }
               
        {$_ -Eq "24" -or $_ -Match "dispute.*logPW"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_pw.ps1; Read-Host 'success!`n`npress enter to exit'";
        }

        {$_ -Eq "25" -or $_ -Match "dispute.*logRAT"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_rat.ps1; Read-Host 'success!`n`npress enter to exit'";
        }

        {$_ -Eq "26" -or $_ -Match "dispute.*logAD"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_ad.ps1; Read-Host 'success!`n`npress enter to exit'";
        }

        {$_ -Eq "27" -or $_ -Match "dispute.*logSL"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_sl.ps1; Read-Host 'success!`n`npress enter to exit'";
        }
        
        {$_ -Eq "28" -or $_ -Match "dispute.*logTAX"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_tax.ps1; Read-Host 'success!`n`npress enter to exit'";
        }

        {$_ -Eq "29" -or $_ -Match "dispute.*logTS"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_TS.ps1; Read-Host 'success!`n`npress enter to exit'";
        }
         
        {$_ -Eq "30" -or $_ -Match "dispute.*logVT"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_vt.ps1; Read-Host 'success!`n`npress enter to exit'";
        }

        {$_ -Eq "31" -or $_ -Match "dispute.*logGT"}
        {
            Start Powershell "$Global:CATButtonsDir\dispute_log_gt.ps1; Read-Host 'success!`n`npress enter to exit'";
        }
        
        {$_ -Eq "32" -or $_ -Match "multi" -or $_ -Eq "M"}
        {
            $Global:MenuStart = Read-Host "Start Num ";
            if ($Global:MenuStart -eq $null -or [string]::IsNullOrEmpty($Global:MenuStart) -or $Global:MenuStart -lt 1)
            {
                $Global:MenuStart = 0;
                $Global:MenuItem = 0;
                run($false);
            }

            $Global:MenuEnd = Read-Host "End Num ";
            if ($Global:MenuEnd -eq $null -or [string]::IsNullOrEmpty($Global:MenuEnd) -or $Global:MenuEnd -gt 98)
            {
                $Global:MenuStart = 0;
                $Global:MenuEnd = 0;
                $Global:MenuItem = 0;
                run($false);
            }

            $Global:MenuItem = $Global:MenuStart;
            Do
            {
                $val=run($false);
                $Global:MenuStart++;
                $Global:MenuItem = $Global:MenuStart;
            }
            While($Global:MenuStart -lt $Global:MenuEnd -and $val -ne "^C")
            If($val -ne "^C") { Write-Host "Calm down..."; }

            $Global:MenuStart = 0;
            $Global:MenuEnd = 0;
        }

        {$_ -Eq "33" -or $_ -Match "cancel"}
        {
            Return '^C';
        }

        {$_ -Eq "99" -or $_ -Match "exit"}
        {
            Exit
        }

        Default
        {
	    if ($Global:MenuStart -eq $null -or [string]::IsNullOrEmpty($Global:MenuStart) -or $Global:MenuStart -lt 1)
            {
                $Global:MenuItem = 0;
                run($false)
            }
        }
    }
    
}

FUNCTION Vacuum-Online
{
    $Filepath = (join-path $Global:AdminDir "VACUUM_ONLINE.TXT");

    if ([System.IO.File]::Exists($Filepath))
    {
        $data = $(cat -delimiter "," -path $Filepath);

        if (((Get-Date) - [datetime]$data[1]).minutes -lt 10)
        {
            return $true;
        }
        else
        {
            remove-item $Filepath;
        }
    }

    return $false;
}
# $($(Get-WmiObject Win32_BIOS | Select-Object SerialNumber).SerialNumber
FUNCTION Remote-Command($Command, $Message)
{
    if (Vacuum-Online)
    {
        $RandName = ([char[]]([char]65..[char]90) + ([char[]]([char]97..[char]122)) + 0..9 | sort {Get-Random})[0..8] -join '';
        write-host $message;
        Add-Content (join-path $Global:AdminDir "REMOTE_COMMAND_$RandName.txt") "$Command";

        for ($i = 0; $i -lt 21; $i++)
        {
            if (Remote-Response($Command))
            {
                return;
            }

            Start-Sleep -Seconds 1;
        }

        Write-Host "Error! wait time for remote response has expired. Please try again later";
    }
    else
    {
        write-host "Error! The Vacuum is currently not running. Please start the vacuum";
    }
}

FUNCTION Remote-Response($Command)
{
    if (Get-Variable 'tmp_*' -Scope Global -ErrorAction 'Ignore')
    {
        Clear-Item Variable:tmp_*;
    }

    $tmp_Files = (Get-ChildItem $Global:AdminDir);
    
    if ($tmp_Files.Count -gt 0)
    {
        ForEach($tmp_File In $tmp_Files.Name)
        {
            if ($tmp_File -like "REMOTE_RETURN_*")
            {
                $tmp_Filepath = Join-Path $Global:AdminDir $tmp_File;
                $tmp_data = (cat -Delimiter ";" -Path $tmp_Filepath);

                
                for ($i=0; $i -lt $tmp_data.count; $i++)
                {
                   $tmp_data[$i] = $tmp_data[$i].remove(($tmp_data[$i].length-1),1);
                }

                If ([string]::IsNullOrEmpty($tmp_Data) -or $tmp_Data.Count -ne 3)
                {
                    if ([System.IO.File]::Exists($tmp_Filepath))
                    {
                        Remove-Item $tmp_Filepath;
                    }
                }
                elseif ($tmp_data[0] -eq $Command)
                {                   
                    if ([System.IO.File]::Exists($tmp_Filepath))
                    {
                        Remove-Item $tmp_Filepath;
                    }
                    write-host "`n";
                    write-host $($tmp_data[2]);

                    return $true;
                }
            }
        }
    }
    return $false;
}

Run -show_menu $true
