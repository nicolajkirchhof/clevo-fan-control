$scriptpath = (split-path $MyInvocation.MyCommand.Path)

Add-Type -LiteralPath "$scriptpath\GetCoreTempInfoNET.dll"
Add-Type -AssemblyName PresentationFramework

$ct = [GetCoreTempInfoNET.CoreTempInfo]::new()

$clevo = get-wmiobject -query "select * from CLEVO_GET" -namespace "root\WMI"


$steps = @(
    (125,  80,  75,  70,  65,  60,  55,  50, 45, 0),
    (255, 230, 200, 180, 155, 135, 128, 100, 64, 0)
)
$cpuLevel = 0
$gpuLevel = 0
$cpuTemp = 0
$gpuTemp = 0
$result = $clevo.SetFanAutoDuty(0)

try{
  while($true){
      do{
          sleep 1.5
          if ($ct.GetData()){
              $newCpuTemp = ($ct.GetTemp[0..3] | measure -Maximum).Maximum
              $newGpuTemp = (get-wmiobject -query "SELECT * FROM Sensor WHERE InstanceId='3850'" -namespace "root\OpenHardwareMonitor").Value
          } else {
              $clevo.SetFanAutoDuty(1);
              [System.Windows.MessageBox]::Show('Clevo Fan Control could not get temperature values, please restart.');
              exit;
          }
      } while([math]::Abs($cpuTemp - $newCpuTemp) -le 3 -and [math]::Abs($gpuTemp - $newGpuTemp) -le 3)
      Write-Output "$(Get-Date -Format o): newCpuTemp $newCpuTemp°C, temp $cpuTemp °C"
      Write-Output "$(Get-Date -Format o): newGpuTemp $newGpuTemp°C, temp $gpuTemp °C"
      $cpuTemp = $newCpuTemp
      $gpuTemp = $newGpuTemp
      $newCpuLevel = 0;
      for(; $steps[0][$newCpuLevel] -gt $cpuTemp; $newCpuLevel++){}
      Write-Output "$(Get-Date -Format o): New Cpu Level $newCpuLevel, Level $cpuLevel"
      
      if ($newCpuLevel -lt $cpuLevel -or $newCpuLevel -gt $cpuLevel+1){
          Write-Output "$(Get-Date -Format o): Setting fan duty to $([int]($steps[1][$newCpuLevel]/255*100))% ($($steps[1][$newCpuLevel]))"
          $cpuLevel = $newCpuLevel;
      }

       $newGpuLevel = 0;
      for(; $steps[0][$newGpuLevel] -gt $gpuTemp; $newGpuLevel++){}
      Write-Output "$(Get-Date -Format o): New Gpu Level $newGpuLevel, Level $gpuLevel"
      
      if ($newGpuLevel -lt $gpuLevel -or $newGpuLevel -gt $gpuLevel+1){
          Write-Output "$(Get-Date -Format o): Setting fan duty to $([int]($steps[1][$newGpuLevel]/255*100))% ($($steps[1][$newGpuLevel]))"
          $gpuLevel = $newGpuLevel;
      }
      $fanValue = ($steps[1][$newGpuLevel] * [uint32]"0x10000") + ($steps[1][$newGpuLevel] * [uint32]"0x0100") + $steps[1][$newCpuLevel]
      Write-Output "New Fan value $("{0:x}" -f $fanValue)"
      $result = $clevo.SetFanDuty($fanValue);
      if ($result.Data1 -eq 104){
          Write-Output "$(Get-Date -Format o): Successful applied new fan duty "
      } else{
          Write-Output "$(Get-Date -Format o): Failure applying new fan duty reverting to automatic"
          $clevo.SetFanAutoDuty(1);
          [System.Windows.MessageBox]::Show('Clevo Fan Control could not apply duty cycle, please restart.');
          exit
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