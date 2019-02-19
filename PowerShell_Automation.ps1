# Created by Sam Escolas in 2016. Edited & Upgraded by Kevin Russell & Shane Johnson on 1/24/2019

#Global Directory Variables
$Global:SourceDir = $(Split-Path (split-path $SCRIPT:MyInvocation.MyCommand.Path -parent) -Parent);

$Global:AdminDir = "$Global:SourceDir\01_Updates\05_Admin\";
$Global:BMI_Filepath = "$Global:SourceDir\01_Updates\01_BMI\";
$Global:PCI_Filepath = "$Global:SourceDir\01_Updates\02_PCI\";
$Global:EventLogDir = "$Global:SourceDir\02_Event_Log\";
$Global:SourceCodeDir = "$Global:SourceDir\03_Source_Code\";

#These variables count the number of updates made in a given session. These are unncessesary
$Global:Maps=0;
$Global:Disputes=0;
$Global:Other=0;
$Global:Errors=0;

#Import the claim class definition for use creating disputes
Import-Module "$Global:SourceCodeDir\Vacuum_Module.psm1";

#Import settings from Vacuum_Settings.xml file into Global variables
Import-Settings

# Error Action Preference Setting (PLZ DO NOT CHANGE)
# This will ask user whether script should continue if it runs into an error
# This will prevent script from looping constantly
$ErrorActionPreference = "Inquire";

# This is C# code that is injected into Powershell to handle Powershell exiting event
# This code will remove the Vacuum_Online.txt in the Admin folder to show that Vacuum is offline
# This code will also write to the Event Log stating that the Vacuum is shutdown
$code = @"
         using System;
         using System.Runtime.InteropServices;
         using System.Management.Automation;
         using System.Management.Automation.Runspaces;
         using System.IO;
         using System.Text;
         
         namespace MyNamespace
         {
             public static class MyClass
             {
                 // public static Runspace defaultRunSpace;
                 private static HandlerRoutine s_rou;

                //  public static void SetHandler(Runspace defRunSpace)
                 public static void SetHandler()
                 {
                     // defaultRunSpace = defRunSpace;
                     if (s_rou == null)
                     {
                        s_rou = new HandlerRoutine(ConsoleCtrlCheck);
                        SetConsoleCtrlHandler(s_rou, true);
                     }
                 }

                 private static bool ConsoleCtrlCheck(CtrlTypes ctrlType)
                 {
                     switch (ctrlType)
                     {    
                         case CtrlTypes.CTRL_CLOSE_EVENT:
                             string path = @"$(join-path $Global:AdminDir "VACUUM_ONLINE.TXT")";
                             string path2 = @"$(join-path $Global:EventLogDir $(get-date).tostring("yyyy-MM-dd")) Event_Log.txt";
                             try
                             {
                                 if (File.Exists(path))
                                 {
                                    File.Delete(path);
                                 }

                                 if (File.Exists(path2))
                                 {
                                    File.AppendAllText(path2,string.Format("{0}{1}", DateTime.Now.ToString(@"hh:mm:ss tt") + " - ", Environment.NewLine));
                                    File.AppendAllText(path2,string.Format("{0}{1}", DateTime.Now.ToString(@"hh:mm:ss tt") + " - The Vacuum closed on $((Get-WmiObject -Class Win32_ComputerSystem -Property Name).Name)'s machine", Environment.NewLine));
                                 }
                                 else
                                 {
                                     using (FileStream fs = File.Create(path2))
                                     {
                                        Byte[] info = new UTF8Encoding(true).GetBytes(DateTime.Now.ToString(@"hh:mm:ss tt") + " - The Vacuum closed on $((Get-WmiObject -Class Win32_ComputerSystem -Property Name).Name)'s machine");
                                        fs.Write(info, 0, info.Length);
                                     }
                                 }
                             }
                             catch (Exception ex)
                             {
                                Console.WriteLine(ex.ToString());
                             }
                             break;
                     }
                     return true;
                 }
          
                 [DllImport("Kernel32")]
                 public static extern bool SetConsoleCtrlHandler(HandlerRoutine Handler, bool Add);
          
                 // A delegate type to be used as the handler routine
                 // for SetConsoleCtrlHandler.
                 public delegate bool HandlerRoutine(CtrlTypes CtrlType);
          
                 // An enumerated type for the control messages
                 // sent to the handler routine.
                 public enum CtrlTypes
                 {
                     CTRL_C_EVENT = 0,
                     CTRL_BREAK_EVENT,
                     CTRL_CLOSE_EVENT,
                     CTRL_LOGOFF_EVENT = 5,
                     CTRL_SHUTDOWN_EVENT
                 }
             }
         }
"@

# This is to load C# code above into a Powershell event handler

$text = Add-Type  -TypeDefinition $code -Language CSharp
#$rs = [System.Management.Automation.Runspaces.Runspace]::DefaultRunspace
#[MyNamespace.MyClass]::SetHandler($rs)
[MyNamespace.MyClass]::SetHandler()

<#
# This is redundancy code to replace the C# event handler if C# code breaks.
# This code works for Powershell exiting of ISO Powershell only.
# If this is to be utilized, you will need to recode all 'Start-Process' commands in CAT_Menu.ps1 and Vacuum_Automation.ps1 to launch Vacuum into ISO instead of PowerShell window
Register-EngineEvent PowerShell.Exiting –Action {
    if ([System.IO.File]::Exists((join-path $Global:AdminDir "VACUUM_ONLINE.TXT")))
    {
        Remove-Item $(join-path $Global:AdminDir "VACUUM_ONLINE.TXT");
    }

    $Global:EventList += "$($(get-date).tostring("hh:mm:ss tt")) - `n";
    $Global:EventList += "$($(get-date).tostring("hh:mm:ss tt")) - The Vacuum closed on $((Get-WmiObject -Class Win32_ComputerSystem -Property Name).Name)'s machine";
    Process_Log;
}
#>

# This class is to hold information for CNR workbook updates. Each update is processed one by one
Class Update
{
    [string] $Source_TBL;
    [int]$Source_ID;
    [string]$Tag;
    [string]$gs_srvID;
    [string]$gs_srvType;
    [string]$Audit_Result;
    [datetime]$Start_Date;
    [string]$Invoice_Date;
    [string]$Dispute_Reason;
    [string]$PON;
    [string]$Claim_Channel;
    [string]$Confidence;
    [string]$USI;
    [string]$Rep;
    [string]$Comment;
    [string]$Edit_Date;
    [string]$Table;
    [string]$Filepath;
    [boolean]$Valid;
    [string]$Problem;
    [array]$Items;
    
    Update() {}

    #construct updates using their filepaths. filepaths themselves contain necessary information
    #files are csvs created by cnr workbooks
    Update([string]$Source_TBL, [string]$Filepath)
    {
        #first we split the file's contents at commas
        $Data = (cat $Filepath) -split ",";
        $this.Filepath=$Filepath;
        $this.Source_TBL = $Source_TBL;

        write-log "Updating file at $Filepath";

        If([string]::IsNullOrEmpty($Data) -or $Data.Count -lt 5) 
        {
            $this.Valid=$false;
            $this.Table="Error";
            $this.Source_ID=0;
            $this.Tag="Error";
            write-log "Error! Data was not found in file";
        }
        Else
        {
            #next we look at the folder the file has been placed in and assign the appropriate data to the update's properties
            Switch(Split-Path (Split-Path $Filepath -Parent) -Leaf)
            {
                01_Map
                {
                    $this.Audit_Result="Mapped";
                    Try {$this.Source_ID=$Data[0];} Catch {$this.Source_ID=0;}
                    $this.Tag="Unmapped";
                    $this.gs_srvType=$Data[1];
                    $this.gs_srvID=$Data[2];
                    $this.Rep=$Data[3];
                    $this.Comment=$Data[4];
                    $this.Edit_Date=$Data[5];
                }

                02_Dispute
                {
                    $this.Audit_Result="Dispute Review";
                    Try {$this.Source_ID=$Data[0];} Catch {$this.Source_ID=0;}
                    $this.Tag=$Data[1];
                    if ($Data[2] -eq $null -or [string]::IsNullOrEmpty($Data[2])) { $this.gStartDate(); } else { $this.Start_Date=eomonth -date $Data[2] -months -1; };
                    $this.Dispute_Reason=$Data[3];
                    if ($Source_TBL -eq "PCI") { $this.Claim_Channel="Email" } else { $this.Claim_Channel=$Data[4] };
                    $this.Confidence=$Data[5];
                    $this.Rep=$Data[6];
                    $this.Comment=$Data[7];
                    $this.Edit_Date=$Data[8];
                    $this.Invoice_Date=$Data[9];
                    $this.USI=$Data[10];
                    $this.PON=$Data[11];
                }

                #the nature of "other" is such that it applies both to the unmapped table and the zero revenue table so we store the table in the name of the file
                05_Other
                {
                    If($filepath -match "x_unmapped_x" -or $filepath -match "x_pciunmapped_x")
                    {
                        Try {$this.Source_ID=$Data[0];} Catch {$this.Source_ID=0;}
                        $this.Tag=$Data[1];
                        $this.Audit_Result=$Data[2];
                        $this.Rep=$Data[3];
                        $this.Comment=$Data[4];
                        $this.Edit_Date=$Data[5];
                        if ($this.Source_TBL -eq "BMI") { $this.Table=$Global:t_Unmapped; } Else { $this.Table=$Global:t_PCI_Unmapped; }
                    }
                    ElseIf($filepath -match "x_cnr_x" -or $filepath -match "x_pcicnr_x")
                    {
                        Try {$this.Source_ID=$Data[0];} Catch {$this.Source_ID=0;}
                        $this.Invoice_Date=$Data[1];
                        $this.Tag=$Data[2];
                        $this.Audit_Result=$Data[3];
                        $this.Rep=$Data[4];
                        $this.Comment=$Data[5];
                        $this.Edit_Date=$Data[6];
                        if ($this.Source_TBL -eq "BMI") { $this.Table=$Global:t_ZeroRevenue } Else { $this.Table=$Global:t_PCI_ZeroRevenue; }
                    }
                }
            }

            If($this.Tag -eq "Unmapped")
            {
                if ($this.Source_TBL -eq "BMI") { $this.Table=$Global:t_Unmapped; } Else { $this.Table=$Global:t_PCI_Unmapped; }
            }
            Else
            {
                if ($this.Source_TBL -eq "BMI") { $this.Table=$Global:t_ZeroRevenue } Else { $this.Table=$Global:t_PCI_ZeroRevenue; }
            }

            $this.Valid=$false;
            $this.Sanitize();
        }
    }

    #prior to making updates we must ensure that they are updates we would like to make
    [boolean]Validate()
    {
        If($this.Source_ID -Eq 0) { Return $false; }
        #indicates whether or not the given gs_srvID/gs_srvType actually exists
        FUNCTION Check-Mapping($gs_srvID, $gs_srvType)
        {
            $table = Switch($gs_srvType)
            {
                "LL" { "ORDER_LOCAL_NEW WHERE ORD_WTN"; }
                "BRD" { "ORDER_BROADBAND WHERE ORD_BRD_ID"; }
                "DED" { "ORDER_DEDICATED WHERE CUS_DED_ID"; }
                "LD" { "ORDER_1PLUS WHERE ORD_WTN"; }
                "TF" { "ORDER_800 WHERE ORD_POTS_ANI_BIL"; }
            }

            $sql = "SELECT COUNT(*) FROM CustomerInventory.$table = '$gs_srvID'";
            Return ((Query($sql))[0] -gt 0);
        }

        #innocent until proven guilty
        $this.Valid=$true;

        #guilty if no cost exists to be disputed or no item exists to map it to
        If($this.Audit_Result -Eq "Mapped") { $this.Valid = (Check-Mapping -gs_srvID $this.gs_srvID -gs_srvType $this.gs_srvType) }
        Return $this.Valid;
    }

    #once we know an update is valid we must apply the update
    [boolean]Upload()
    {
        If($this.Source_ID -Eq 0) { Return $false; }

        #actual represents the number of updates that have taken place during any given interation
        $actual=0;

        #expected defines our expectation. this is dependent on the number of tables an update must affect
        $expected=0;

        $this.Items=@();

	    If ($this.Audit_Result -Eq "Dispute Review")
        {
            write-log "$($this.Rep) is disputing $($this.Source_TBL) $($this.Source_ID) to $($this.Claim_Channel)";

            Create-Disputes -Source_TBL $($this.Source_TBL) -Source_ID $($this.Source_ID) -Start_Date $($this.Start_Date) -Dispute_Reason $($this.Dispute_Reason) -PON $($this.PON) -Dispute_Category 'GRT CNR' -Audit_Type 'CNR Audit' -Rep $($this.Rep) -Source 'CNR Disputes' -Claim_Channel $($this.Claim_Channel) -Confidence $($this.Confidence) -USI $($this.USI);
	    }

        Switch($this.Audit_Result)
        {
            #if we are mapping something we expect to update two tables -- BMM/BMB and Unmapped
            "Mapped"
            {
                If($this.Filepath -match "x_bmb_x") 
                {
                    $expected=3;

                    switch ($this.Source_TBL)
                    {
                        "BMI"
                        {
                            Write-log "$($this.Rep) is updating $($this.Source_TBL) $($this.Source_ID) to $($this.Audit_Result) on $($Global:t_Unmapped), $($Global:t_BMM), and $($Global:t_BMB)";
                            
                            $this.Items += Query("declare @TMP table (Action varchar(10), TBL varchar(255), ID int, gs_SrvType varchar(3), gs_SrvID varchar(50), frameID varchar(10), Edit_Date datetime); INSERT INTO $($Global:t_BMB) ( gs_srvID, gs_srvType, BMI_ID, Rep ) OUTPUT 'INSERT' As Action, '$($Global:t_BMB)' As TBL, INSERTED.ID, NULL gs_srvType, NULL gs_srvID, NULL frameID, NULL Edit_Date INTO @TMP VALUES ('$($this.gs_srvID)', '$($this.gs_srvType)', '$($this.Source_ID)', '$($this.Rep)'); select * from @TMP");

                            if ((Query("WITH M ( id, dt ) AS (SELECT BMI_ID, MAX(Edit_Date) FROM $($Global:t_Unmapped) GROUP BY BMI_ID) SELECT Audit_Result FROM $($Global:t_Unmapped) CNR INNER JOIN M ON M.id=CNR.BMI_ID AND M.dt=CNR.Edit_Date WHERE BMI_ID = $($this.Source_ID)")).Audit_Result -Eq "Mapped")
                            {
                                $actual++;
                            }
                            else
                            {
                                $this.Items += Query("declare @TMP table (Action varchar(10), TBL varchar(255), ID int, gs_SrvType varchar(3), gs_SrvID varchar(50), frameID varchar(10), Edit_Date datetime); INSERT INTO $($Global:t_Unmapped) (BMI_ID, Audit_Result, Rep, Comment, Edit_Date) OUTPUT 'INSERT' As Action, '$($Global:t_Unmapped)' As TBL, INSERTED.ID, NULL gs_srvType, NULL gs_srvID, NULL frameID, NULL Edit_Date INTO @TMP VALUES ('$($this.Source_ID)', '$($this.Audit_Result)', '$($this.Rep)', '$($this.Comment)', '$([string]([datetime]::Parse($this.Edit_Date)).AddSeconds(-17))'); select * from @TMP");
                            }

                            $this.Items += Query("declare @TMP table (Action varchar(10), TBL varchar(255), BMM_ID int, gs_SrvType varchar(3), gs_SrvID varchar(50), frameID varchar(10), Edit_Date datetime); UPDATE $($Global:t_BMM) SET gs_srvID='$($this.Source_ID)', gs_srvType='$([string]("BMB"))', frameID='$($this.Rep)', Edit_Date=getdate() OUTPUT 'UPDATE' As Action, '$($Global:t_BMM)' As TBL, INSERTED.BMM_ID, DELETED.gs_srvType, DELETED.gs_srvID, DELETED.frameID, DELETED.Edit_Date INTO @TMP WHERE BMI_ID='$($this.Source_ID)'; select * from @TMP");
                            
                            $actual += $this.Items.Count;
                        }
                        "PCI"
                        {
                            Write-log "$($this.Rep) is updating $($this.Source_TBL) $($this.Source_ID) to $($this.Audit_Result) on $($Global:t_PCI_Unmapped),  $($Global:t_PCI), and $($Global:t_PCI_BMB)";

                            $this.Items += Query("declare @TMP table (Action varchar(10), TBL varchar(255), ID int, gs_SrvType varchar(3), gs_SrvID varchar(50), frameID varchar(10), Edit_Date datetime); INSERT INTO $($Global:t_PCI_BMB) (gs_srvID, gs_srvType, PCI_ID, Rep) OUTPUT 'INSERT' As Action, '$($Global:t_PCI_BMB)' As TBL, INSERTED.ID, NULL gs_srvType, NULL gs_srvID, NULL frameID, NULL Edit_Date INTO @TMP VALUES ('$($this.gs_srvID)', '$($this.gs_srvType)', '$($this.Source_ID)', '$($this.Rep)'); select * from @TMP");

                            if ((Query("WITH M ( id, dt ) AS (SELECT PCI_ID, MAX(Edit_Date) FROM $($Global:t_PCI_Unmapped) GROUP BY PCI_ID) SELECT Audit_Result FROM $($Global:t_PCI_Unmapped) CNR INNER JOIN M ON M.id=CNR.PCI_ID AND M.dt=CNR.Edit_Date WHERE PCI_ID = $($this.Source_ID)")).Audit_Result -Eq "Mapped")
                            {
                                $actual++;
                            }
                            else
                            {
                                $this.Items += Query("declare @TMP table (Action varchar(10), TBL varchar(255), ID int, gs_SrvType varchar(3), gs_SrvID varchar(50), frameID varchar(10), Edit_Date datetime); INSERT INTO $($Global:t_PCI_Unmapped) (PCI_ID, Audit_Result, Rep, Comment, Edit_Date) OUTPUT 'INSERT' As Action, '$($Global:t_PCI_Unmapped)' As TBL, INSERTED.ID, NULL gs_srvType, NULL gs_srvID, NULL frameID, NULL Edit_Date INTO @TMP VALUES ('$($this.Source_ID)', '$($this.Audit_Result)', '$($this.Rep)', '$($this.Comment)', '$([string]([datetime]::Parse($this.Edit_Date)).AddSeconds(-17))'); select * from @TMP");
                            }

                            $this.Items += Query("declare @TMP table (Action varchar(10), TBL varchar(255), ID int, gs_SrvType varchar(3), gs_SrvID varchar(50), frameID varchar(10), Edit_Date datetime); UPDATE $($Global:t_PCI) SET gs_srvID='$($this.gs_srvID)', gs_srvType='$($this.gs_srvType)', Edit_Date=getdate() OUTPUT 'UPDATE' As Action, '$($Global:t_PCI)' As TBL, INSERTED.ID, DELETED.gs_srvType, DELETED.gs_srvID, NULL frameID, DELETED.Edit_Date INTO @TMP WHERE ID='$($this.Source_ID)'; select * from @TMP");
                            
                            $actual += $this.Items.Count;
                        }
                    }
                }
                Else
                {
                    $expected=2;

                    switch ($this.Source_TBL)
                    {
                        "BMI"
                        {

                            Write-log "$($this.Rep) is updating $($this.Source_TBL) $($this.Source_ID) to $($this.Audit_Result) on $($Global:t_Unmapped) and $($Global:t_BMM)";

                            $this.Items += Query("declare @TMP table (Action varchar(10), TBL varchar(255), BMM_ID int, gs_SrvType varchar(3), gs_SrvID varchar(50), frameID varchar(10), Edit_Date datetime); UPDATE $($Global:t_BMM) SET gs_srvID='$($this.gs_srvID)', gs_srvType='$($this.gs_srvType)', frameID='$($this.Rep)', Edit_Date=getdate() OUTPUT 'UPDATE' As Action, '$($Global:t_BMM)' As TBL, INSERTED.BMM_ID, DELETED.gs_srvType, DELETED.gs_srvID, DELETED.frameID, DELETED.Edit_Date INTO @TMP WHERE BMI_ID='$($this.Source_ID)'; select * from @TMP");
                            $this.Items += Query("declare @TMP table (Action varchar(10), TBL varchar(255), ID int, gs_SrvType varchar(3), gs_SrvID varchar(50), frameID varchar(10), Edit_Date datetime); INSERT INTO $($Global:t_Unmapped) (BMI_ID, Audit_Result, Rep, Comment, Edit_Date) OUTPUT 'INSERT' As Action, '$($Global:t_Unmapped)' As TBL, INSERTED.ID, NULL gs_srvType, NULL gs_srvID, NULL frameID, NULL Edit_Date INTO @TMP VALUES ('$($this.Source_ID)', '$($this.Audit_Result)', '$($this.Rep)', '$($this.Comment)', '$([string]([datetime]::Parse($this.Edit_Date)).AddSeconds(-17))'); select * from @TMP");

                            $actual += $this.Items.Count;
                        }
                        "PCI"
                        {
                            Write-log "$($this.Rep) is updating $($this.Source_TBL) $($this.Source_ID) to $($this.Audit_Result) on $($Global:t_PCI_Unmapped) and  $($Global:t_PCI)";

                            $this.Items += Query("declare @TMP table (Action varchar(10), TBL varchar(255), ID int, gs_SrvType varchar(3), gs_SrvID varchar(50), frameID varchar(10), Edit_Date datetime); UPDATE $($Global:t_PCI) SET gs_srvID='$($this.gs_srvID)', gs_srvType='$($this.gs_srvType)', Edit_Date=getdate() OUTPUT 'UPDATE' As Action, '$($Global:t_PCI)' As TBL, INSERTED.ID, DELETED.gs_srvType, DELETED.gs_srvID, NULL frameID, DELETED.Edit_Date INTO @TMP WHERE ID='$($this.Source_ID)'; select * from @TMP");
                            $this.Items += Query("declare @TMP table (Action varchar(10), TBL varchar(255), ID int, gs_SrvType varchar(3), gs_SrvID varchar(50), frameID varchar(10), Edit_Date datetime); INSERT INTO $($Global:t_PCI_Unmapped) (PCI_ID, Audit_Result, Rep, Comment, Edit_Date) OUTPUT 'INSERT' As Action, '$($Global:t_PCI_Unmapped)' As TBL, INSERTED.ID, NULL gs_srvType, NULL gs_srvID, NULL frameID, NULL Edit_Date INTO @TMP VALUES ('$($this.Source_ID)', '$($this.Audit_Result)', '$($this.Rep)', '$($this.Comment)', '$([string]([datetime]::Parse($this.Edit_Date)).AddSeconds(-17))'); select * from @TMP");

                            $actual += $this.Items.Count;
                        }
                    }
                }

                If($actual -eq $expected)
                {
                    $Global:Maps++;
                }
                elseif ($this.Items.Count -gt 0)
                {
                    Write-Correction-Log -List $this.items;
                }
            }

            #otherwise we only expect to update the unmapped/zero revenue table
            Default
            {
                if (-not $Global:DErrors)
                {
                    $expected=1;

                    Write-log "$($this.Rep) is updating $($this.Source_TBL) $($this.Source_ID) to '$($this.Audit_Result)' on $($this.Table)";

                    switch ($this.Source_TBL)
                    {
                        "BMI"
                        {
                            If($this.Table -Eq $Global:t_Unmapped)
                            {
                                $this.Items += Query("declare @TMP table (Action varchar(10), TBL varchar(255), ID int, gs_SrvType varchar(3), gs_SrvID varchar(50), frameID varchar(10), Edit_Date datetime); INSERT INTO $($this.Table) (BMI_ID, Audit_Result, Rep, Comment, Edit_Date) OUTPUT 'INSERT' As Action, '$($this.Table)' As TBL, INSERTED.ID, NULL gs_srvType, NULL gs_srvID, NULL frameID, NULL Edit_Date INTO @TMP VALUES ('$($this.Source_ID)', '$($this.Audit_Result)', '$($this.Rep)', '$($this.Comment)', '$($this.Edit_Date)'); select * from @TMP");
                            }
                            Else
                            {
                                $this.Items += Query("declare @TMP table (Action varchar(10), TBL varchar(255), ID int, gs_SrvType varchar(3), gs_SrvID varchar(50), frameID varchar(10), Edit_Date datetime); INSERT INTO $($this.Table) (BMI_ID, Invoice_Date, Tag, Audit_Result, Rep, Comment, Edit_Date) OUTPUT 'INSERT' As Action, '$($this.Table)' As TBL, INSERTED.ID, NULL gs_srvType, NULL gs_srvID, NULL frameID, NULL Edit_Date INTO @TMP VALUES ('$($this.Source_ID)', '$($this.Invoice_Date)', '$($this.Tag)', '$($this.Audit_Result)', '$($this.Rep)', '$($this.Comment)', '$($this.Edit_Date)'); select * from @TMP");
                            }

                            $actual += $this.Items.Count;
                        }
                        "PCI"
                        {
                            If($this.Table -Eq $Global:t_PCI_Unmapped)
                            {
                                $this.Items += Query("declare @TMP table (Action varchar(10), TBL varchar(255), ID int, gs_SrvType varchar(3), gs_SrvID varchar(50), frameID varchar(10), Edit_Date datetime); INSERT INTO $($this.Table) (PCI_ID, Audit_Result, Rep, Comment, Edit_Date) OUTPUT 'INSERT' As Action, '$($this.Table)' As TBL, INSERTED.ID, NULL gs_srvType, NULL gs_srvID, NULL frameID, NULL Edit_Date INTO @TMP VALUES ('$($this.Source_ID)', '$($this.Audit_Result)', '$($this.Rep)', '$($this.Comment)', '$($this.Edit_Date)'); select * from @TMP");
                            }
                            Else
                            {
                                $this.Items += Query("declare @TMP table (Action varchar(10), TBL varchar(255), ID int, gs_SrvType varchar(3), gs_SrvID varchar(50), frameID varchar(10), Edit_Date datetime); INSERT INTO $($this.Table) (PCI_ID, Invoice_Date, Tag, Audit_Result, Rep, Comment, Edit_Date) OUTPUT 'INSERT' As Action, '$($this.Table)' As TBL, INSERTED.ID, NULL gs_srvType, NULL gs_srvID, NULL frameID, NULL Edit_Date INTO @TMP VALUES ('$($this.Source_ID)', '$($this.Invoice_Date)', '$($this.Tag)', '$($this.Audit_Result)', '$($this.Rep)', '$($this.Comment)', '$($this.Edit_Date)'); select * from @TMP");
                            }

                            $actual += $this.Items.Count;
                        }
                    }
                
                    If($actual -eq $expected)
                    {
                        If($this.Audit_Result -Eq "Dispute Review")
                        {
                            $Global:Disputes++;
                        }
                        Else
                        {
                            $Global:Other++;
                        }
                    }
                    elseif ($this.Items.Count -gt 0)
                    {
                        Write-Correction-Log -List $this.items;
                    }
                }
            }
        }

        #indicate whether or not the expected number of updates were made
        Return ($actual -Eq $expected);
    }

    Hidden Error()
    {
        #if something goes wrong, this method deals with moving the csv where elsewhere
        $subfolder=($this.Filepath.split("\\"))[-2];
        Move-Item $this.Filepath (Join-Path (Split-Path (Split-Path $this.Filepath -Parent) -Parent) "06_Errors\$subfolder\$((Split-Path $this.Filepath -Leaf))");

        $Global:Errors++;
    }

    Hidden Delete()
    {
        #if an item is successfully updated, this method removes it from the queue
        if ([System.IO.File]::Exists($this.Filepath))
        {
            Remove-Item $this.Filepath;
        }
    }

    Hidden Sanitize()
    {
        #this is necessary to ensure that text columns don't contain single quotes -- they will break SQL
        $this.Comment=$this.Comment -replace "'", "''";
    }

    Hidden gStartDate()
    {
        #if no start date is provided in the CNR workbooks upon dispute, we maximize the amount of cose we dispute based on the dispute limitations
        switch ($this.Source_TBL)
        {
            "BMI"
            {
                $sql = "SELECT eomonth(dateadd(month, -1, MIN(INV.Bill_Date))) [Start_Date]
                FROM $($Global:t_BMI) BMI INNER JOIN $($Global:t_Limit) LIM ON LIM.BAN = BMI.BAN INNER JOIN $($Global:t_InvSummary) INV ON INV.BAN = BMI.BAN
                WHERE BMI_ID = '$($this.Source_ID)' AND INV.Bill_Date > DATEADD(D, (-1 * LIM.Dispute_Limit) + 7 , GETDATE())";
                $this.Start_Date = (Query($sql)).Start_Date;
            }
            
            "PCI"
            {
                $sql = "SELECT eomonth(dateadd(month, -1, MIN(INV.Bill_Date))) [Start_Date]
                FROM $($Global:t_PCI) PCI INNER JOIN $Global:t_Limit LIM ON LIM.BAN = PCI.BAN INNER JOIN $($Global:t_PC) INV ON INV.BAN = PCI.BAN
                WHERE PCI_ID = '$($this.Source_ID)' AND INV.Bill_Date > DATEADD(D, (-1 * LIM.Dispute_Limit) + 7 , GETDATE())";
                $this.Start_Date = (Query($sql)).Start_Date;
            }
        }
    }
}

# This is code for Remote commands that is triggered by the CAT_Menu.ps1
# This will make host of Vacuum restart/Stop Vacuum or request the status of the Vacuum
FUNCTION Admin-Remote
{
    if (Get-Variable 'Files' -Scope Global -ErrorAction 'Ignore')
    {
        Clear-Item Variable:Files;
    }
    if (Get-Variable 'File' -Scope Global -ErrorAction 'Ignore')
    {
        Clear-Item Variable:File;
    }
    if (Get-Variable 'Filepath' -Scope Global -ErrorAction 'Ignore')
    {
        Clear-Item Variable:Filepath;
    }
    if (Get-Variable 'Remote_Name' -Scope Global -ErrorAction 'Ignore')
    {
        Clear-Item Variable:Auth_Name;
    }
    if (Get-Variable 'CurrentTime' -Scope Global -ErrorAction 'Ignore')
    {
        Clear-Item Variable:CurrentTime;
    }

    $Files = (Get-ChildItem $Global:AdminDir);
    if ($Files.Count -gt 0)
    {
        ForEach($File In $Files.Name)
        {
            if ($File -like "REMOTE_COMMAND_*")
            {
                $Filepath = Join-Path $Global:AdminDir $File;
                $data = cat $Filepath;
                $Remote_Name = $(Get-Acl $Filepath).Owner;
                if ([System.IO.File]::Exists($Filepath))
                {
                    Remove-Item $Filepath;
                }

                If ([string]::IsNullOrEmpty($Data))
                {
                    next;
                }
                else
                {
                    $RandName = ([char[]]([char]65..[char]90) + ([char[]]([char]97..[char]122)) + 0..9 | sort {Get-Random})[0..8] -join '';
                    switch ($data)
                    {
                        "Restart"
                        {
                            write-log "Vacuum Automation is restarting per ADMIN, $Remote_Name, request...";
                            Process_Log;

                            Start-Process powershell.exe "& '$Global:SourceCodeDir\Vacuum_Automation.ps1'" -WindowStyle Minimized;

                            $Global:EventList += "$data;1;Vacuum restarted on $((Get-WmiObject -Class Win32_ComputerSystem -Property Name).Name)'s machine"
                            Process_Log2 -Filename "REMOTE_RETURN_$RandName.TXT";

                            return $true;
                        }
                        "Shutdown"
                        {
                            write-log "Vacuum Automation is stopping per ADMIN, $Remote_Name, request...";
                            Process_Log;

                            $Global:EventList += "$data;1;Vacuum stopped on $((Get-WmiObject -Class Win32_ComputerSystem -Property Name).Name)'s machine"
                            Process_Log2 -Filename "REMOTE_RETURN_$RandName.TXT";

                            return $true;
                        }
                        "Status"
                        {
                            $CurrentTime = ((get-date)-$startTime);

                            write-log "Vacuum Automation retrieved Vacuum status per ADMIN $Remote_Name";
                            write-log "`n";
                            Process_Log

                            $Global:EventList += "$data;1;Vacuum is currently running on $((Get-WmiObject -Class Win32_ComputerSystem -Property Name).Name)'s machine for $([string]::Format("`r{0:d2} Days ({1:d2}:{2:d2}:{3:d2})", $CurrentTime.Days, $CurrentTime.hours, $CurrentTime.minutes, $CurrentTime.seconds)) and has processed the following items:";
                            $Global:EventList += "Mapped: $($Global:Maps)  Disputed: $($Global:Disputes)  Other: $($Global:Other)  Errors: $($Global:Errors)";
                            Process_Log2 -Filename "REMOTE_RETURN_$RandName.TXT";

                            break;
                        }
                        default
                        {
                            $Global:EventList += "$data;0;Error! Remote command $data does is not a valid command"
                            Process_Log2 -Filename "REMOTE_RETURN_$RandName.TXT";
                        }
                    }
                }
            }
        }
    }

    return $false;
}

# This checks appropriate folders for updates
FUNCTION Updates-Exist
{
    if (Get-Variable 'Folders' -Scope Global -ErrorAction 'Ignore')
    {
        Clear-Item Variable:Folders;
    }

    [array]$Folders = "01_Map", "02_Dispute", "05_Other";

    foreach ($Folder in $Folders)
    {
        if ((Get-ChildItem (Join-Path $Global:BMI_Filepath $Folder)).Count -gt 0)
        {
            return @("BMI", (Join-Path $Global:BMI_Filepath $Folder));
        }
        elseif ((Get-ChildItem (Join-Path $Global:PCI_Filepath $Folder)).Count -gt 0)
        {
            return @("PCI", (Join-Path $Global:PCI_Filepath $Folder));
        }
    }
    return $null;
}

# This processes updates if an update is found
FUNCTION Process-Updates($MyArr)
{
    $Path = $MyArr[1];

    write-log "Processing updates in $Path"
    $Updates = @();

    ForEach($Update In (Get-ChildItem $Path).Name)
    {
        $Updates+=(([update]::New($MyArr[0], (Join-Path $Path $Update))));
    }

    $Updates | ForEach {
        if ($_.Audit_Result -eq "Mapped")
        {
            write-log "Checking data validation for $($_.Source_TBL) $($_.Source_ID) (Gs_Serv $($_.gs_srvType) $($_.gs_srvID)) by $($_.Rep)"
        }
        If($_.Validate())
        {
            if ($_.Audit_Result -eq "Mapped")
            {
                write-log "Item $($_.Source_TBL) $($_.Source_ID) by $($_.Rep) passed data validation";
            }

            $Global:DErrors = $false;

            If(($_.Upload()))
            {
                if ($Global:DErrors)
                {
                    write-log "Item $($_.Source_TBL) $($_.Source_ID) by $($_.Rep) was not disputed. Failed Dispute";
                    write-log "No updates have been made";
                    mv $_.Filepath $_.Filepath.Replace(".txt", "__nodispute.txt");
                    $_.Filepath = $_.Filepath.Replace(".txt", "__nodispute.txt");
                    $_.Error();
                }
                else
		        {
                    write-log "Item $($_.Source_TBL) $($_.Source_ID) by $($_.Rep) updated to '$($_.Audit_Result)' on $($_.Table)";
                    $_.Delete();
                }
            }
            Else
            {
                write-log "Item $($_.Source_TBL) $($_.Source_ID) by $($_.Rep) failed to be updated.";
                mv $_.Filepath $_.Filepath.Replace(".txt", "__upload_error.txt");
                $_.Filepath = $_.Filepath.Replace(".txt", "__upload_error.txt");
                $_.Error();
            }
        }
        Else
        {
            write-log "Item $($_.Source_TBL) $($_.Source_ID) by $($_.Rep) failed data validation. GS_Srv $($_.gs_srvType) $($_.gs_srvID) does not exist";
            mv $_.Filepath $_.Filepath.Replace(".txt", "__invalid.txt");
            $_.Filepath = $_.Filepath.Replace(".txt", "__invalid.txt");
            $_.Error();
        }
    }
}

[boolean] $My_Continue = $true;
$Global:CorrList = @();
$Global:EventList = @();
$StartTime = (Get-Date);

write-host "`n`n`n";
write-log "`n";

write-log "Starting up the Vacuum Automation on $((Get-WmiObject -Class Win32_ComputerSystem -Property Name).Name)'s machine...";
Process_Log

BOT_Online -timestamp $StartTime;

try
{
    Do
    {
        Write-Progress -Activity "Mapped: $($Global:Maps)  Disputed: $($Global:Disputes)  Other: $($Global:Other)  Errors: $($Global:Errors;)" -Status "Waiting" -Id 1;

        While(($Path = Updates-Exist) -Eq $null)
        {
            if (-not $My_Continue)
            {
                $My_Continue = $true;
                write-log "Vacuum Automation waiting for more updates...";
                Process_Log;
            }

            if (Admin-Remote)
            {
                Exit
            }

            if (((Get-Date) - $Global:LastRefresh).minutes -gt 4)
            {
                BOT_Online -timestamp (Get-Date);
            }

            Start-Sleep -Seconds 1;
        }

        Process-Updates -MyArr $Path;

        if (Admin-Remote)
        {
            Exit
        }

        if (((Get-Date) - $Global:LastRefresh).minutes -gt 4)
        {
            BOT_Online -timestamp (Get-Date);
        }

        write-log "`n";
        $My_Continue = $false;
    }
    Until(1 -eq 2)
}
catch
{
    Quit-PS -message "Error! $($Error[0].Exception.Message). The Vacuum exited on $((Get-WmiObject -Class Win32_ComputerSystem -Property Name).Name)'s machine...";

    Read-Host "Press Enter to exit..." | Out-Null;
}
finally
{
    Quit-PS -message "The Vacuum closed on $((Get-WmiObject -Class Win32_ComputerSystem -Property Name).Name)'s machine";
}

Exit
