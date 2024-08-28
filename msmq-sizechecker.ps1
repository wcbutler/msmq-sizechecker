#8/28/2024 WCB

Clear-Host
$currentTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Output "MSMQ logging started : $currentTime" 

# Define the machine names
$machines = @("machinename1","machinename2","machinename3")

# Define the queue name if needed
$queuename = ""


# Function to get and display the MSMQ queue sizes, CPU, and memory usage
function Get-MSMQInfo {
    # Initialize an array to store the results
    $results = @()

    # Get the current time
    $currentTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Loop through each machine and get the required information
    foreach ($machine in $machines) {
        try {
            # Get the list of MSMQ queues on the machine
            $queues = Get-WmiObject -Namespace "root\cimv2" -Class "Win32_PerfFormattedData_MSMQ_MSMQQueue" -ComputerName $machine | Where-Object { $_.Name -like "*$queuename*" }

            # Get CPU and memory usage
            $cpu = Get-WmiObject -Namespace "root\cimv2" -Class "Win32_Processor" -ComputerName $machine | Measure-Object -Property LoadPercentage -Average | Select-Object -ExpandProperty Average
            $cpu = [math]::Round($cpu)
            $memory = Get-WmiObject -Namespace "root\cimv2" -Class "Win32_OperatingSystem" -ComputerName $machine
            $totalMemory = [math]::round($memory.TotalVisibleMemorySize / 1MB, 2)
            $freeMemory = [math]::round($memory.FreePhysicalMemory / 1MB, 2)
           # $usedMemory = $totalMemory - $freeMemory

            # Loop through each queue and store the results
            foreach ($queue in $queues) {
                $highQueueSize = $queue.MessagesInQueue -gt 50
                $results += [pscustomobject]@{
                    Machine        = $machine
                    Queue          = $queue.Name
                    Messages       = $queue.MessagesInQueue
                    CPU_Percentage = $cpu
                   # Used_Memory_MB = $usedMemory
                    Free_Memory_MB = $freeMemory
                    HighQueueSize  = $highQueueSize
                }
            }
        } catch {
            Write-Output "Failed to retrieve data from $machine. Error: $_"
        }
    }

    # Display the current time
    Write-Output "Current Time: $currentTime"

    # Display the results in a table
    if ($results.Count -gt 0) {
        $results | Format-Table -AutoSize
    } else {
        Write-Output "No queues found with names like 'archiver*' on the specified machines."
    }
}

# Run the function initially and then every 30 seconds
while ($true) {
    Get-MSMQInfo
    Start-Sleep -Seconds 30
}