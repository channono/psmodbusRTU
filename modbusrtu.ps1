# modbus RTU for Powershell by channono
[console]::TreatControlCAsInput = $true
while ($true) {
    if ($Host.UI.RawUI.KeyAvailable -and (3 -eq [int]$Host.UI.RawUI.ReadKey("AllowCtrlC,IncludeKeyUp,NoEcho").Character)) {
        break
        [System.GC]::Collect()
     # 用户按下了 Ctrl+C，执行清理操作
    }

# 配置串口
if ($IsLinux){
    $portName = "/dev/ttyS0"  # 替换为实际的串口名称
}
if ($IsMacOS) {
    #$portName = "/dev/cu.usbserial-00000000"  # 替换为实际的串口名称
    $portName = "/dev/tty.usbserial-00000000"  # 替换为实际的串口名称
}
if ($IsWindows) {
    $portName = "com3"  # 替换为实际的串口名称
}
$baudRate = 9600  # Replace with the desired baud rate
$parity = "None"  # Replace with the desired parity (e.g., "None", "Even", "Odd")
$dataBits = 8  # Replace with the desired number of data bits
$stopBits = 1  # Replace with the desired number of stop bits (e.g., "One", "OnePointFive", "Two")

$port = New-Object System.IO.Ports.SerialPort($portName, $baudRate, $parity, $dataBits, $stopBits)

$port.Open()

# 读取 Holding Register（功能码 03）
function ReadHoldingRegister($slaveAddress, $startAddress, $registerCount) {

    $functionCode = [byte]3
    $request = [byte[]]@(
        [byte]$slaveAddress,
        $functionCode,
        [System.Convert]::ToByte($startAddress / 256),
        [System.Convert]::ToByte($startAddress % 256),

        [System.Convert]::ToByte($registerCount / 256),
        [System.Convert]::ToByte($registerCount % 256)
    )
   # Write-Host "RAW data: $request"
    $request += CalculateCRC  $request 
     
    Write-Host "sended data: $request"
    $port.Write($request, 0, $request.Length)

    # 等待数据返回
    Start-Sleep -Milliseconds 100
    $response = $port.ReadExisting()
    $responseBytes = [System.Text.Encoding]::ASCII.GetBytes($response)
    if ($responseBytes) {
        Write-Host "Received bytes:"
        $responseBytes -join " " | Write-Host  

        return $responseBytes
    }else {
        Write-Host "No response received"
    }
    
}
 

# 计算 CRC
function CalculateCRC($data) {    
    $crc = 0xFFFF 

    foreach ($byte in $data) {
        $crc = $crc -bxor $byte

        for ($i = 0; $i -lt 8; $i++) {
            if ($crc -band 1) {
                $crc = ($crc -shr 1) -bxor 0xA001
            } else {
                $crc = $crc -shr 1
            }
        }
    }
   # Write-Host  $crc
    $crcBytes = [System.BitConverter]::GetBytes($crc)
   # [Array]::Reverse($crcBytes)
    return $crcBytes[0..1]      
}
function GetValue($responseBytes){
        # Display the bytes read from holding registers
        Write-Host "Bytes read from holding registers:"
        $holdingRegisterBytes = $responseBytes[3..($responseBytes.Length - 3)]
        #$holdingRegisterBytes | ForEach-Object { Write-Host $_ }
        $holdingRegisterBytes -join " " | Write-Host
        Write-Host "Values read from holding registers:"
        for ($i = 0; $i -lt $holdingRegisterBytes.Length-1; $i += 4) {
            $valueBytes = $holdingRegisterBytes[$i..($i + 3)]
            [Array]::Reverse($valueBytes)  # Reverse the order of bytes
            $intValue = [BitConverter]::ToInt32($valueBytes, 0)
            $actualValue = $intValue * 0.1
            Write-Host "Actual value: $actualValue"

        }
}

    try {
        #ReadHoldingRegister 2 23296 51
        while ($true) {
            $responseBytes =     ReadHoldingRegister 2 23298 6 
            GetValue $responseBytes   
            # 在这里处理读取到的数据，例如打印、保存到文件等
            # ...
        # write-host $holdingRegisters
            # 等待一段时间后再次读取
            #Start-Sleep -Seconds 1
            Start-Sleep -Milliseconds 3000
        }
    }finally {
        # 清理操作，例如关闭串口连接
        # 关闭串口
        $port.Close()

    }
}








