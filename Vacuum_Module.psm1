# Created by Sam Escolas in 2016. Edited & Upgraded by Kevin Russell & Shane Johnson on 1/24/2019

# Global Variables (Plz dont change)
[string]$Global:Server;
[string]$Global:Database;
[string]$Global:HadoopDSN;
[string]$Global:t_BMM;
[string]$Global:t_BMB;
[string]$Global:t_Unmapped;
[string]$Global:t_ZeroRevenue;
[string]$Global:t_DisputeStaging;
[string]$Global:t_EmailExtract;
[string]$Global:t_PCI;
[string]$Global:t_PCI_BMB;
[string]$Global:t_PCI_Unmapped;
[string]$Global:t_PCI_ZeroRevenue;
[string]$Global:t_PC;
[string]$Global:t_MRC;
[string]$Global:t_MRCCMP;
[string]$Global:t_OCC;
[string]$Global:t_BMI;
[string]$Global:t_Limit;
[string]$Global:t_InvSummary;
[boolean]$Global:DErrors;
[array]$Global:EventList;
[array]$Global:CorrList;
$Global:LastRefresh;

# This will query SQL Server of SQL statements
	# This is easier on the eyes than Invoke-Sqlcmd with the same server/database repeated each time

FUNCTION Query
{
    param
    (
        [Parameter(Mandatory=$true)]$SQL,
        [Parameter(Mandatory=$false)][boolean]$Retry
    )

    $error.clear();

    Try
    {
        $Data = Invoke-Sqlcmd -ServerInstance $Global:Server -Database $Global:Database -QueryTimeout 0 -Query $SQL -ErrorAction SilentlyContinue -OutputSqlErrors $false;
    }
    Catch
    {
        if ($Retry)
        {
            write-log "Error! Failed to query $Global:HadoopDSN connection. Writing to error log";
            write-log "$Global:Server -> $sql";
            Return $null;
        }
        else
        {
            write-log "Warning! $Global:Settings_Server Query failed. Retrying same query in 5 seconds";
            Start-Sleep -Seconds 5;
            write-log "Re-attempting query again...";
            Return (Query -SQL $SQL -Retry $true);
        }
    }

    If ($Data -Eq $null -or [string]::IsNullOrEmpty($Data) -or $Data.Count -Eq 0)
    {
        Return $null;
    }
    else
    {
        Return $Data;
    }
}

# This will query Hadoop Cluster of SQL statements
	# Computer must have Cloudera Impala ODBC driver installed with ODBC connection created in Data Sources under Administrative Tools of Control Panel
    # If Hadoop Cluster fails, revert back to using ODS solely for the proccess of MRC/OCC seeds. See code below in functions where code tagged as (REPLACEMENT CODE) is noted
    # If implementing replacement code, please comment out code tagged as (HADOOP DEPENDENT)

# (HADOOP DEPENDENT START)

function Hquery
{
    param
    (
        [Parameter(Mandatory=$true)]$SQL,
        [Parameter(Mandatory=$false)][boolean]$Retry
    )

    $error.clear();

    $conn = "DSN=$Global:HadoopDSN;DATABASE=default;Trusted_Connection=Yes;";
    $data = New-Object System.Data.DataSet;

    Try
    {
        (New-Object System.Data.Odbc.OdbcDataAdapter($SQL, $conn)).Fill($data) | out-null;
    }
    Catch
    {
        if ($Retry)
        {
            write-log "Error! Failed to query $Global:HadoopDSN connection. Writing to error log";
            write-log "$Global:HadoopDSN -> $sql";
            Return $null;
        }
        else
        {
            write-log "Warning! $Global:HadoopDSN Query failed. Retrying same query in 5 seconds";
            Start-Sleep -Seconds 5;
            write-log "Re-attempting query again...";
            Return $(Hquery -SQL $SQL -Retry $true);
        }
    }

    if ($Data.Tables[0].Rows.Count -eq 0)
    {
        Return $null;
    }
    else
    {
        Return $Data;
    }
}

# (HADOOP DEPENDENT END)


#This executes SQL commands and will return the number of rows affected by a single transaction
FUNCTION Execute($SQL)  
{
    Try
    {
        Return (Query("$SQL SELECT @@ROWCOUNT"))[0];
    }
    Catch
    {
        Return -1;
    }
}

Function Import-Settings
{
    [xml]$ConfigFile = get-content (join-path $Global:SourceCodeDir "Vacuum_Settings.xml");

    $Global:Server = $ConfigFile.Settings.Network.Server;
    $Global:Database = $ConfigFile.Settings.Network.Database;
    $Global:HadoopDSN = $ConfigFile.Settings.Network.HadoopDSN;

    $Global:t_BMM = $ConfigFile.Settings.Read_Write_TBL.BMM;
    $Global:t_BMB = $ConfigFile.Settings.Read_Write_TBL.BMB;
    $Global:t_Unmapped = $ConfigFile.Settings.Read_Write_TBL.Unmapped;
    $Global:t_ZeroRevenue = $ConfigFile.Settings.Read_Write_TBL.ZeroRevenue;
    $Global:t_DisputeStaging = $ConfigFile.Settings.Read_Write_TBL.DisputeStaging;
    $Global:t_EmailExtract = $ConfigFile.Settings.Read_Write_TBL.EmailExtract;
    $Global:t_PCI = $ConfigFile.Settings.Read_Write_TBL.PCI;
    $Global:t_PCI_BMB = $ConfigFile.Settings.Read_Write_TBL.PCI_BMB;
    $Global:t_PCI_Unmapped = $ConfigFile.Settings.Read_Write_TBL.PCI_Unmapped;
    $Global:t_PCI_ZeroRevenue = $ConfigFile.Settings.Read_Write_TBL.PCI_ZeroRevenue;

    $Global:t_PC = $ConfigFile.Settings.Read_TBL.PaperCost;
    $Global:t_MRC = $ConfigFile.Settings.Read_TBL.MRC;
    $Global:t_MRCCMP = $ConfigFile.Settings.Read_TBL.MRC_CMP;
    $Global:t_OCC = $ConfigFile.Settings.Read_TBL.OCC;
    $Global:t_BMI = $ConfigFile.Settings.Read_TBL.BMI;
    $Global:t_Limit = $ConfigFile.Settings.Read_TBL.Limitations;
    $Global:t_InvSummary = $ConfigFile.Settings.Read_TBL.Invoice_Summary;
}

FUNCTION Quit-PS($message)
{
    if ([System.IO.File]::Exists((join-path $Global:AdminDir "VACUUM_ONLINE.TXT")))
    {
        Remove-Item $(join-path $Global:AdminDir "VACUUM_ONLINE.TXT");
    }

    $Global:EventList += "$($(get-date).tostring("hh:mm:ss tt")) - `n";
    $Global:EventList += "$($(get-date).tostring("hh:mm:ss tt")) - $message";

    Process_Log;
}

Function Write-Log([string]$Log_TXT)
{
    $Global:EventList += "$($(get-date).tostring("hh:mm:ss tt")) - $Log_TXT"

    write-host "$($(get-date).tostring("(MM/dd/yyyy) hh:mm:ss tt")) - $Log_TXT";
}

function Write-Correction-Log($List)
{
    $Global:CorrList += "$($(get-date).tostring("hh:mm:ss tt")) - `n";
    $Global:CorrList += "$($(get-date).tostring("hh:mm:ss tt")) - ------> Need to remove/fix records:";
    $Global:CorrList += "$($(get-date).tostring("hh:mm:ss tt")) - (Action, Action_TBL, Action_ID, Prev_Gs_SrvType, Prev_Gs_SrvID, Prev_FrameID, Prev_Edit_Date)";

    foreach ($line in $List)
    {
        $Global:CorrList += "$($(get-date).tostring("hh:mm:ss tt")) - ('$($line.ItemArray -join "', '")'),";
    }
}

function Process_Log()
{
    if ($Global:CorrList.count -gt 0)
    {
        Add-Content $(join-path $Global:EventLogDir "$($(get-date).tostring("yyyy-MM-dd")) Correction_Log.txt") $Global:CorrList;
        $Global:CorrList = @();
    }

    if ($Global:EventList.count -gt 0)
    {
         Add-Content $(join-path $Global:EventLogDir "$($(get-date).tostring("yyyy-MM-dd")) Event_Log.txt") $Global:EventList;
        $Global:EventList = @();
    }
}

function Process_Log2($Filename)
{
    if ($Global:EventList.count -gt 0)
    {
        Add-Content $(join-path $Global:AdminDir $Filename) $Global:EventList;
        $Global:EventList = @();
    }
}

function BOT_Online([datetime]$timestamp)
{
    Process_Log

    if ([System.IO.File]::Exists((join-path $Global:AdminDir "VACUUM_ONLINE.TXT")))
    {
        Remove-Item (join-path $Global:AdminDir "VACUUM_ONLINE.TXT");
    }

    $Global:EventList += "$((Get-WmiObject -Class Win32_ComputerSystem -Property Name).Name),$timestamp";
    Process_Log2 -Filename "VACUUM_ONLINE.TXT";

    $Global:LastRefresh = $timestamp;
}

#Grabs the next batch date in integer format (yyyyMMdd)
FUNCTION Get-NextBatch()
{
    $today = Get-Date
    For($i = 0; $i -le 7; $i++)
    {
        If($today.DayOfWeek -eq "Friday"){$friday = $today;} Else{$today = $today.AddDays(1);}
    }   
    
    If ($friday.Month -le 9){$month = "0" + $friday.Month} Else {$month = $friday.Month}
    If ($friday.Day -le 9){$day = "0" + $friday.Day} Else {$day = $friday.Day}
    
    Return [string]$friday.Year + [string]$month + [string]$day
}

# Finds the MRC/OCC BDT seeds according to BMI ID
	# Script checks BMI table for necessary information and grabs seeds accordingly
    # ODS Replacement for Get-Seeds is located at the bottom of the script

# (HADOOP DEPENDENT START)
FUNCTION Get-Seeds($Source_TBL, $Source_ID, $Start_Date)
{
    switch ($Source_TBL)
    {
        "BMI"
        {
            $Data = Query("SELECT Vendor, BAN, BTN, WTN, Circuit_ID FROM $($Global:t_BMI) WHERE BMI_ID = $Source_ID");

	        $sql = "select 'MRC' As Src_TBL, bdt_mrc_id As Src_ID from $Global:t_MRCCMP where BMI_ID = $Source_ID and Invoice_Date > '$($start_date.tostring("yyyy-MM-dd"))'";
            $sql += " union all select upper(Activity_Type) As Src_TBL, bdt_occ_id As Src_ID from $Global:t_OCC where Invoice_Date > '$($start_date.tostring("yyyy-MM-dd"))' AND Vendor = '$($Data.Vendor)' AND BAN = '$($Data.BAN)' AND Amount > 0 AND isnull(BTN, '') = '$($Data.WTN)' AND ISNULL(Circuit_ID, '') = '$($Data.Circuit_ID)'";
        }

        "PCI"
        {
            $sql = "select 'MRC' As Src_TBL, Seed As Src_ID from $Global:t_PC where PCI_ID = $Source_ID and ((trunc(Bill_Date, 'MM') + interval 1 month) - interval 1 day) > '$($start_date.tostring("yyyy-MM-dd"))' and MRC > 0"
            $sql += " union all select 'NRC' As Src_TBL, Seed As Src_ID from $Global:t_PC where PCI_ID = $Source_ID and ((trunc(Bill_Date, 'MM') + interval 1 month) - interval 1 day) > '$($start_date.tostring("yyyy-MM-dd"))' and NRC > 0"
            $sql += " union all select 'FRAC' As Src_TBL, Seed As Src_ID from $Global:t_PC where PCI_ID = $Source_ID and ((trunc(Bill_Date, 'MM') + interval 1 month) - interval 1 day) > '$($start_date.tostring("yyyy-MM-dd"))' and FRAC > 0"
        }
    }

    Return HQuery -SQL $sql;
}
# (HADOOP DEPENDENT END)

# This function is referenced by Process_BMI_Updates.ps1 script (BMI ID Support Only)
	#This will Grab BDT seeds and insert necessary data in a class called claim
	#Function will go through the claim upload process so that new disputes can show in Staging or Email
FUNCTION Create-Disputes($Source_TBL, $Source_ID, $Start_Date, $Dispute_Reason, $PON, $Dispute_Category, $Audit_Type, $Rep, $Source, $Claim_Channel, $Confidence, $USI)
{
    $Baseline = (Query("SELECT (SELECT COUNT(*) FROM $($Global:t_DisputeStaging)) + (SELECT COUNT(*) FROM $($Global:t_EmailExtract)) [ct]")).ct;
    
    
    #utilizes the claim class to create and file disputes
    write-log "Searching for $Source_TBL cost for $Source_TBL $Source_ID from $($start_date.tostring("yyyy-MM-dd")) to now";

    # (HADOOP DEPENDENT START)
    $Seeds = Get-Seeds -Source_TBL $Source_TBL -Source_ID $Source_ID -Start_Date $Start_Date;
    # (HADOOP DEPENDENT END)

    <#
        # (REPLACEMENT CODE START)
        $Seeds = Get-Seeds-ODS -BMI_ID $BMI_ID -Start_Date $Start_Date;
        # (REPLACEMENT CODE END)
    #>
    
    if ($Seeds -ne $null)
    {
        $i=0;

        # (HADOOP DEPENDENT START)
        $total=$Seeds.tables[0].Rows.Count
        # (HADOOP DEPENDENT END)

        <#
            # (REPLACEMENT CODE START)
            $total=$Seeds.Count
            # (REPLACEMENT CODE END)
        #>

        if ($total -gt 0)
        {
            write-log "Found $total seeds of cost for $Source_TBL $Source_ID";

            $Claims = @();

            # (HADOOP DEPENDENT START) #
            if ($seeds -ne $null)
            {
                foreach ($Seed in $Seeds.tables[0].rows)
                {
                    # Add-Content $(join-path $Global:EventLogDir "Seed_OUTPUT.TXT") "$($Seed[0]), $($Seed[1])"
                    [array]$Claims+=[claim]::New($Source_TBL, $($Seed[1]), $($Seed[0]), $Dispute_Reason, $Dispute_Category, $Audit_Type, $Rep, $Source, $Claim_Channel, $Confidence, 0, $USI, $PON) 
                };
            }
            # (HADOOP DEPENDENT END) #

            <#
                # (REPLACEMENT CODE START)
                $Seeds | ForEach
                {
                    [array]$Claims+=[claim]::New($_.Seed, $_.Table, $Dispute_Reason, $Dispute_Category, $Audit_Type, $Rep, $Source, $Claim_Channel, $Confidence, 0, $USI, $PON) 
                };
                # (REPLACEMENT CODE END)
            #>

            $i=0;
            $mrc_error=0;
            $occ_error=0;

            $Claims | ForEach {

                Write-Progress -Activity "Uploading dispute $i of $total" -PercentComplete (100 * ($i++/$total)) -Id 2;

                if ($_.Error -ne 1)
                {
                    If(-not (([string]::IsNullOrEmpty($_.Seed)) -or ([string]::IsNullOrEmpty($_.Batch)) -or ([string]::IsNullOrEmpty($_.Record_Type))))
                    {
                        $_.STC_Claim_Number = "$($_.Batch)_$($_.Record_Type.Substring(0,1).ToUpper())$($_.Seed)";
                    }
                    $_.Upload() | Out-Null;
                }
                else
                {
                    if ($_.Record_Type = "MRC")
                    {
                        $mrc_error += 1
                    }
                    else
                    {
                        $occ_error += 1
                    }
                }
            }

            Write-Progress -completed $true -id 2;
            if ($mrc_error -gt 0 -or $occ_error -gt 0)
            {
                write-log "Notice - Wasn't able to find cost details for $mrc_error MRC and $occ_error OCC seeds out of $total seeds";
            }
        }
        else
        {
            write-log "Wasn't able to find any cost for $Source_TBL $Source_ID";
        }
    }

    If (((Query("SELECT (SELECT COUNT(*) FROM $($Global:t_DisputeStaging)) + (SELECT COUNT(*) FROM $($Global:t_EmailExtract)) [ct]")).ct - $Baseline) -Eq 0)
    {
        $Global:DErrors = $true;
    }
    else
    {
        write-log "$Source_TBL $Source_ID has been disputed to $Claim_Channel";
    }
}

# Returns the end of the month (works the same as the SQL function
FUNCTION EOMONTH([datetime]$Date, [int]$Months=0)
{
    $Date=$Date.AddMonths($Months);
    Return [datetime]::Parse("$($Date.Month)/$((($Date.AddMonths(1)).AddDays(-1*$Date.Day)).Day)/$($Date.Year)");
}

# This Class is named Claim and this class stores necessary data for the creation of new claims
	# This class will go format the data to SQL format and insert into the appropriate tables
Class Claim
{
    #first we list the properties of the class with their respective data types
    [string]$Source_TBL;
    [string]$Vendor;
    [string]$Platform;
    [string]$Dispute_Category;
    [string]$STC_Claim_Number;
    [string]$Record_Type;
    [string]$BAN;
    [datetime]$Bill_Date;
    [double]$Billed_Amt;
    [double]$Claimed_Amt;
    [string]$Dispute_Reason;
    [string]$USI;
    [string]$USOC;
    [string]$BTN;
    [string]$WTN;
    [string]$Circuit_ID;
    [string]$PON;
    [string]$CLLI;
    [string]$Usage_Rate;
    [string]$MOU;
    [string]$Jurisdiction;
    [string]$Sender_Email;
    [string]$Short_Paid;
    [int]$Batch=(Get-NextBatch);
    [string]$Comment
    [string]$Audit_Type;
    [string]$Claim_Channel;
    [string]$Confidence;
    [string]$Display_Status;
    [string]$Ilec_Confirmation;
    [string]$Ilec_Comment;
    [string]$Causing_SO;
    [string]$Escalate;
    [string]$Close_Reason;
    [string]$Norm_Close_Reason;
    [float]$Approved_Amount;
    [float]$Received_Amount;
    [string]$Rep;
    [string]$Credit_Received_Invoice_Date;
    [string]$Source;
    [string]$Phrase_Code;
    [int]$Seed;
    [boolean]$SDN;
    [boolean]$Error;

    #then we create class constructors 
    Claim() {} #allows you to create an empty claim using [Claim]::New() or New-Object Claim

    Claim($Source_TBL, $Seed, $Record_Type, $Dispute_Reason, $Dispute_Category, $Audit_Type, $Rep, $Source, $Claim_Channel, $Confidence, $Claimed_Amt, $USI, $PON)
    {
        $this.Source_TBL = $Source_TBL;
        $this.Seed=$Seed;
        $this.PON=$PON;
        $this.Record_Type=$Record_Type;
        If($PON -Eq $Null -Or $PON -Eq "") { $this.bSeed($Source_TBL, $False); } Else { $this.bSeed($Source_TBL, $True); }
        $this.Dispute_Reason=$Dispute_Reason;
        $this.Dispute_Category=$Dispute_Category;
        $this.Audit_Type=$Audit_Type;
        $this.Rep=$Rep;
        $this.Source=$Source;
        $this.Claim_Channel=$Claim_Channel;
        $this.Confidence=$Confidence;
        If($Claimed_Amt -Eq 0) { $this.Claimed_Amt=$this.Billed_Amt; } Else { $this.Claimed_Amt=$Claimed_Amt; }
        If($this.USI -Eq "" -Or $this.USI -Eq $Null) { $this.bUSI(); } Else { $this.USI=$USI; }
    }

    [int]Upload()
    {
        If($this.Claim_Channel -Ne "STC") 
        { 
            $this.Display_Status="Filed";
            Return $this.Upload("email"); 
        } 
        Else
        { 
            Return $this.Upload("staging");
        }
    }

    [int]Upload($Table)
    {
        FUNCTION Generate-SQL($Table)
        {
            FUNCTION Generate-Values([int]$Table)
            {
                FUNCTION Append-String([string]$Value, [string]$SQL, [boolean]$First_Item)
                {
                    Try
                    {
                        If($First_Item)
                        {
                            If([string]::IsNullOrEmpty($Value))
                            {
                                $t_sql = $SQL + "NULL";
                            }
                            Else
                            {
                                $t_sql = $SQL + "'" + [string]$Value.Replace("'", "''") + "'";
                            }
                        }
                        Else
                        {
                            If([string]::IsNullOrEmpty($Value))
                            {
                                $t_sql = $SQL + ", NULL";
                            }
                            Else
                            {
                                $t_sql = $SQL + ", '" + [string]$Value.Replace("'", "''") + "'";
                            }
                        }
                        Return $t_sql;
                    }
                    Catch
                    {
                        If($First_Item)
                        {
                            If([string]::IsNullOrEmpty($Value))
                            {
                                $t_sql = $SQL + "NULL";
                            }
                            Else
                            {
                                $t_sql = $SQL + "'" + [string]$Value + "'";
                            }
                        }
                        Else
                        {
                            If([string]::IsNullOrEmpty($Value))
                            {
                                $t_sql = $SQL + ", NULL";
                            }
                            Else
                            {
                                $t_sql = $SQL + ", '" + [string]$Value + "'";
                            }
                        }
                        Return $t_sql;
                    }
                }
        
                FUNCTION Append-Date([datetime]$Value, [string]$SQL, [boolean]$First_Item)
                {
                    If($First_Item)
                    {
                        If([string]::IsNullOrEmpty($Value) -or ($Value.Year -le 1990))
                        {
                            $t_sql = $SQL + "NULL";
                        }
                        Else
                        {
                            $t_sql = $SQL + "'" + $Value + "'";
                        }
                    }
                    Else
                    {
                        If([string]::IsNullOrEmpty($Value) -or ($Value.Year -le 1928))
                        {
                            $t_sql = $SQL + ", NULL";
                        }
                        Else
                        {
                            $t_sql = $SQL + ", '" + $Value + "'";
                        }
                    }
                    Return $t_sql;
                }

                FUNCTION Append-Number([double]$Value, [string]$SQL, [boolean]$First_Item)                                                                                            
                {
                    If($First_Item)
                    {
                        If([string]::IsNullOrEmpty($Value))
                        {
                            $t_sql = $SQL + "NULL";
                        }
                        Else
                        {
                            $t_sql = $SQL + "" + $Value;
                        }
                    }
                    Else
                    {
                        If([string]::IsNullOrEmpty($Value))
                        {
                            $t_sql = $SQL + ", NULL";
                        }
                        Else
                        {
                            $t_sql = $SQL + ", " + $Value;
                        }
                    }
                    Return $t_sql;
                }

                Switch($Table)
                {
                    0 
                    {
                        $sql="";
                        $sql = Append-String -Value $this.Vendor -SQL $sql -First_Item $true;
                        $sql = Append-String -Value $this.Platform -SQL $sql -First_Item $false;
                        $sql = Append-Date -Value $this.Bill_Date -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Display_Status -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.STC_Claim_Number -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.BAN -SQL $sql -First_Item $false;
                        $sql = Append-Date -Value ([datetime]::New($this.Batch.ToString().Substring(0,4),$this.Batch.ToString().Substring(4,2),$this.Batch.ToString().Substring(6,2))) -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.ILEC_Confirmation -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.ILEC_Comment -SQL $sql -First_Item $false;
                        $sql = Append-Number -Value $this.Claimed_Amt -SQL $sql -First_Item $false;
                        $sql = Append-Number -Value $this.Approved_Amount -SQL $sql -First_Item $false;
                        $sql = Append-Number -Value $this.Received_Amount -SQL $sql -First_Item $false;
                        If([string]::IsNullOrEmpty(($this.Credit_Received_Invoice_Date))) { $sql+=", NULL"; } Else { $sql+=", '$($this.Credit_Received_Invoice_Date)'"; }
                        $sql = Append-String -Value $this.Escalate -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Close_Reason -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Audit_Type -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Claim_Channel -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Confidence -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.USI -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Dispute_Reason -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.USOC -SQL $sql -First_Item $false;
                        $sql = Append-Date -Value (Get-Date) -SQL $sql -First_Item $false;
                        $sql = Append-Number -Value $this.Batch -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Source -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Rep -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Dispute_Category -SQL $sql -First_Item $false;
			$sql = Append-String -Value $this.Comment -SQL $sql -First_Item $false;
                    }
                    
                    1
                    {
                        $sql = "";
                        $sql = Append-String -Value $this.Vendor -SQL $sql -First_Item $true;
                        $sql = Append-String -Value $this.Platform -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Dispute_Category -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.STC_Claim_Number -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Record_Type -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.BAN -SQL $sql -First_Item $false;
                        $sql = Append-Date -Value $this.Bill_Date -SQL $sql -First_Item $false;
                        $sql = Append-Number -Value $this.Billed_Amt -SQL $sql -First_Item $false;
                        $sql = Append-Number -Value $this.Claimed_Amt -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Dispute_Reason -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.USI -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.USOC -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Phrase_Code -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Causing_SO -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.PON -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.CLLI -SQL $sql -First_Item $false;
                        If([string]::IsNullOrEmpty(($this.Usage_Rate)))
                        {
                            $sql += ", NULL";
                        }
                        Else
                        {
                            $sql = Append-String -Value ([string]$this.Usage_Rate) -SQL $sql -First_Item $false;
                        }

                        If([string]::IsNullOrEmpty(($this.MOU)))
                        {
                            $sql += ", NULL";
                        }
                        Else
                        {
                            $sql = Append-String -Value ([string]$this.Minutes_Of_Usage) -SQL $sql -First_Item $false;
                        }
                        $sql = Append-String -Value $this.Jurisdiction -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Short_Paid -SQL $sql -First_Item $false;
                        $sql = Append-Number -Value $this.Batch -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Comment -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Audit_Type -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Confidence -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Rep -SQL $sql -First_Item $false;
                        $sql = Append-String -Value $this.Source -SQL $sql -First_Item $false;
                        If($this.Display_Status -notmatch "reject")
                        {
                            $sql += ", NULL, NULL, NULL, '$(Get-Date)'"
                        }
                        Else
                        {
                            $sql += ", 'Rejected', '$($this.Close_Reason)', $($this.Norm_Close_Reason)', '$(Get-Date)'"
                        }
                        
                    }
                }

                Return $sql;
            }

            Switch($Table)
            {
                {$_ -match "email" }
                {
		            $SQL="INSERT INTO ";
		            $SQL += $($Global:t_EmailExtract);
		            $SQL += " (Vendor, Platform, Bill_Date, Display_Status, STC_Claim_Number, BAN, Date_Submitted, ILEC_Confirmation, ILEC_Comments, Dispute_Amount, Credit_Approved, Credit_Received_Amount, Credit_Received_Bill_Date, Escalate, Close_Escalate_Reason, Audit_Type, Claim_Channel, Confidence_Level, USI, Dispute_Reason, USOC, Date_Updated, Batch, Source, Rep, Dispute_Category, Comment)";
                    $SQL += " VALUES (";
		            $SQL += $(Generate-Values(0));
                    $SQL += ")";
                }

                {$_ -match "staging"}
                {
                    $SQL="INSERT INTO ";
		            $SQL += $($Global:t_DisputeStaging);
		            $SQL += " (Vendor, Platform, Dispute_Category, STC_Claim_Number, Record_Type, BAN, Bill_Date, Billed_Amt, Claimed_Amt, Dispute_Reason, USI, USOC, Billed_Phrase_Code, Causing_SO, PON, CLLI, Usage_Rate, MOU, Jurisdiction, Short_Paid, Batch, Comment, Audit_Type, Confidence, Rep, Source, Status, Rejection_Reason, Norm_Rejection_Reason, Edit_Date)";
		            $SQL += " VALUES (";
		            $SQL += $(Generate-Values(1));
		            $SQL += ")";
                }
            }
	        # Add-Content $(join-path $Global:EventLogDir "SQL_OUTPUT2.TXT") $SQL
            Return $SQL;
        }

	Return Execute(Generate-SQL($Table));
    }

    Hidden bUSI()
    {
        If([string]::IsNullOrEmpty($this.WTN) -Or $this.WTN -Eq '0000000000')
        {
            If([string]::IsNullOrEmpty($this.Circuit_ID))
            {
                $this.USI=$this.BTN;
            }
            Else
            {
                $this.USI=$this.Circuit_ID;
            }
        }
        Else
        {
            $this.USI=$this.WTN;
        }
    }

    Hidden bSeed($Source_TBL, [BOOLEAN]$PON)
    #currently only build for MRC and OCC disputes
    {
        $Data=$this.gSeedData($Source_TBL);
        If($Data -ne $null)
        {
            $this.Vendor=$Data.Vendor;
            $this.Platform=$Data.Platform;
            $this.BAN=$Data.BAN;
            $this.Bill_Date=$Data.Bill_Date;
            $this.Billed_Amt=$Data.Amount;
            $this.BTN=$Data.BTN;
            $this.Circuit_ID=$Data.Circuit_ID;
            $this.SDN=$true;
            $this.USOC=$Data.USOC;

            Switch($this.Record_Type)
            {
                "MRC"
                {
                    $this.WTN=$Data.WTN;
                }

                default
                {
                    If(-Not $PON) { $this.PON=$Data.PON; }
                    $this.Causing_SO=$Data.SO;
                }
            }
        }
        else
        {
            $this.error = 1;
        }
    }

    Hidden [object]gSeedData($Source_TBL)
    {
        If([string]::IsNullOrEmpty($this.Seed)) { return $null }

	    $SQL = ""
        switch ($Source_TBL)
        {
            "BMI"
            {
	            switch ($this.Record_Type)
	            {
                    {$_ -match "MRC" }
                    {
		                $SQL = "SELECT BDT_MRC_ID, Vendor, Platform, Bill_Date, State, BAN, BTN, WTN, Circuit_ID, USOC, Amount, INVOICE_DATE, BMI_ID Source_ID, NULL PON, NULL SO FROM $Global:t_MRC WHERE BDT_MRC_ID = $($this.Seed) and amount > 0";
	                }
	                default
	                {
		                $SQL = "SELECT BDT_OCC_ID, Vendor, Platform, Bill_Date, State, BAN, BTN, NULL As WTN, Circuit_ID, USOC, Amount, INVOICE_DATE, NULL Source_ID, PON, SO FROM $Global:t_OCC WHERE BDT_OCC_ID = $($this.Seed) and amount > 0";
	                }
	            }
            }

            "PCI"
            {
                $SQL = "SELECT Seed, Vendor, NULL Platform, Bill_Date, State, BAN, BTN, WTN, Circuit_ID, NULL As USOC, $($this.Record_Type) Amount, eomonth(Bill_Date) INVOICE_DATE, PCI_ID Source_ID, NULL PON, NULL SO FROM $Global:t_PC WHERE Seed = $($this.Seed)";
            }
        }

        #Add-Content $(join-path $Global:EventLogDir "SQL_OUTPUT.TXT") $SQL

        Return Query($SQL);
    }
}

<#
    # (REPLACEMENT CODE START)
    Function Get-Seeds-ODS
    ($BMI_ID, $Start_Date)
    {
            #returns a list of seeds we hope to dispute
            # Old Algorithm to collect Seeds

            $Data=Query("SELECT Vendor, BAN, BTN, WTN, Circuit_ID FROM $($Global:t_BMI) WHERE BMI_ID = $BMI_ID");

            $sql = "SELECT BDT_MRC_ID [Seed], 'MRC' [Table] FROM $($Global:t_MRC)";
            $sql += " WHERE Bill_Date >= '$Start_Date' AND Vendor = '$($Data.Vendor)' AND BAN = '$($Data.BAN)' AND Amount > 0 AND ISNULL(BTN, '') = '$($Data.BTN)' AND ISNULL(WTN, '') = '$($Data.WTN)' AND ISNULL(Circuit_ID, '') = '$($Data.Circuit_ID)'";
        
            $sql += " UNION SELECT BDT_OCC_ID [Seed], 'OCC' [Table] FROM $($Global:t_OCC)";
            $sql += " WHERE Bill_Date >= '$Start_Date' AND Vendor = '$($Data.Vendor)' AND BAN = '$($Data.BAN)' AND Amount > 0 AND isnull(BTN, '') = '$($Data.WTN)' AND ISNULL(Circuit_ID, '') = '$($Data.Circuit_ID)'";
    }
    # (REPLACEMENT CODE END)
#>
