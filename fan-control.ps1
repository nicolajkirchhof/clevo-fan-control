$scriptpath = (split-path $MyInvocation.MyCommand.Path)

Add-Type -LiteralPath "$scriptpath\GetCoreTempInfoNET.dll"
Add-Type -AssemblyName PresentationFramework

$ct = [GetCoreTempInfoNET.CoreTempInfo]::new()

$clevo = get-wmiobject -query "select * from CLEVO_GET" -namespace "root\WMI"


$cpuSteps = @(
    (125,  80,  75,  70,  65,  60,  55,  50, 45, 0),
    (255, 230, 210, 190, 165, 150, 128, 100, 64, 0)
)

$gpuSteps = @(
    (125,  80,  75,  70,  65,  60,  55,  50, 45, 0),
    (255, 230, 200, 180, 165, 150, 100, 86, 64, 0)
)
$cpuLevel = 0
$gpuLevel = 0
$cpuTemp = 0
$gpuTemp = 0
$result = $clevo.SetFanAutoDuty(0)
$fanValue = 0

while($true){
  try{
      do{
          sleep 1.5
          if ($ct.GetData()){
              $newCpuTemp = ($ct.GetTemp[0..3] | measure -Maximum).Maximum
              $gpuTemps = (get-wmiobject -query "SELECT * FROM Sensor WHERE Name='GPU Core' AND SensorType='Temperature'" -namespace "root\OpenHardwareMonitor")
              $newGpuTemp = $gpuTemps.Value
              if (-not $newGpuTemp){
                # if ( $gpuTemps.Max -le 60 ){
                $newGpuTemp = $gpuTemp
                # } else {
                  # $newGpuTemp = $gpuTemps.Max
                # }
                # Write-Output "Could not read GPU-Temperature! Using $newGpuTemp°C"
              }
          } else {
              $clevo.SetFanAutoDuty(1);
              [System.Windows.MessageBox]::Show('Clevo Fan Control could not get temperature values, please restart.');
          }
      } while([math]::Abs($cpuTemp - $newCpuTemp) -le 3 -and [math]::Abs($gpuTemp - $newGpuTemp) -le 3)
      Write-Output "$(Get-Date -Format o): newCpuTemp $newCpuTemp°C, temp $cpuTemp °C"
      Write-Output "$(Get-Date -Format o): newGpuTemp $newGpuTemp°C, temp $gpuTemp °C"
      $cpuTemp = $newCpuTemp
      $gpuTemp = $newGpuTemp
      $newCpuLevel = 0;
      for(; $cpuSteps[0][$newCpuLevel] -gt $cpuTemp; $newCpuLevel++){}
      Write-Output "$(Get-Date -Format o): New Cpu Level $newCpuLevel, Level $cpuLevel"
      
      if ($newCpuLevel -lt $cpuLevel -or $newCpuLevel -gt $cpuLevel+1){
          Write-Output "$(Get-Date -Format o): Setting fan duty to $([int]($cpuSteps[1][$newCpuLevel]/255*100))% ($($cpuSteps[1][$newCpuLevel]))"
          $cpuLevel = $newCpuLevel;
      }

       $newGpuLevel = 0;
      for(; $gpuSteps[0][$newGpuLevel] -gt $gpuTemp; $newGpuLevel++){}
      Write-Output "$(Get-Date -Format o): New Gpu Level $newGpuLevel, Level $gpuLevel"
      
      if ($newGpuLevel -lt $gpuLevel -or $newGpuLevel -gt $gpuLevel+1){
          Write-Output "$(Get-Date -Format o): Setting fan duty to $([int]($gpuSteps[1][$newGpuLevel]/255*100))% ($($gpuSteps[1][$newGpuLevel]))"
          $gpuLevel = $newGpuLevel;
      }
      $newFanValue = ($gpuSteps[1][$gpuLevel] * [uint32]"0x10000") + ($gpuSteps[1][$gpuLevel] * [uint32]"0x0100") + $cpuSteps[1][$cpuLevel]
      Write-Output "New Fan value $("{0:x}" -f $newFanValue)"
      if ($newFanValue -ne $fanValue){
        $fanValue = $newFanValue
        $result = $clevo.SetFanDuty($fanValue);
        if ($result.Data1 -eq 104){
            Write-Output "$(Get-Date -Format o): Successful applied new fan duty"
        } else{
            Write-Output "$(Get-Date -Format o): Failure applying new fan duty reverting to automatic"
            $clevo.SetFanAutoDuty(1);
            [System.Windows.MessageBox]::Show('Clevo Fan Control could not apply duty cycle, please restart.');
        }
      }

  }
  catch {
    $clevo.SetFanAutoDuty(1);
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
     Write-Output "$(Get-Date -Format o): Unknown Failure. Item: $FailedItem, Message: $ErrorMessage"
    [System.Windows.MessageBox]::Show('Clevo Fan Control had an unknown error, please restart'); 
  }
}