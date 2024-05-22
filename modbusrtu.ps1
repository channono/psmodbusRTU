# modbus RTU for Powershell by channono


##------------------------------------------------------------------------------- 
# FC01 Read Coil Status
function ReadCoilStatus($slaveAddress, $startAddress, $coilCount) {
    $functionCode = [byte]1
    $byteCount = [System.Convert]::ToByte([math]::Ceiling($coilCount / 8))

    # Construct the request
    $request = [byte[]]@(
        [byte]$slaveAddress,
        $functionCode,
        [System.Convert]::ToByte($startAddress / 256),
        [System.Convert]::ToByte($startAddress % 256),
        [System.Convert]::ToByte($coilCount / 256),
        [System.Convert]::ToByte($coilCount % 256)
    )

    # Calculate and append CRC16
    $request += CalculateCRC $request
    $request -join " " | Write-Host

    # Send the request to the Modbus device (you need to set up the port first)
    $port.Write($request, 0, $request.Length)

    # Wait for response (adjust sleep time as needed)
    Start-Sleep -Milliseconds 100

    # Read the response (you need to handle the response accordingly)
    $response = $port.ReadExisting()
    $responseBytes = [System.Text.Encoding]::ASCII.GetBytes($response)

    if ($responseBytes) {
        Write-Host "Received bytes:"
        $responseBytes -join " " | Write-Host
        return $responseBytes
    } else {
        Write-Host "No response received"
    }
}

#FC02 Read Discrete Input status
function ReadDiscreteInputs($slaveAddress, $startAddress, $inputCount) {
    $functionCode = [byte]2
    $byteCount = [System.Convert]::ToByte([math]::Ceiling($inputCount / 8))

    # Construct the request
    $request = [byte[]]@(
        [byte]$slaveAddress,
        $functionCode,
        [System.Convert]::ToByte($startAddress / 256),
        [System.Convert]::ToByte($startAddress % 256),
        [System.Convert]::ToByte($inputCount / 256),
        [System.Convert]::ToByte($inputCount % 256)
    )

    # Calculate and append CRC16
    $request += CalculateCRC $request
    $request -join " " | Write-Host

    # Send the request to the Modbus device (you need to set up the port first)
    $port.Write($request, 0, $request.Length)

    # Wait for response (adjust sleep time as needed)
    Start-Sleep -Milliseconds 100

    # Read the response (you need to handle the response accordingly)
    $response = $port.ReadExisting()
    $responseBytes = [System.Text.Encoding]::ASCII.GetBytes($response)

    if ($responseBytes) {
        Write-Host "Received bytes:"
        $responseBytes -join " " | Write-Host
        return $responseBytes
    } else {
        Write-Host "No response received"
    }
}
   
# FC03 Read Holding Register
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

#FC04 Read Input Register
function ReadInputRegister($slaveAddress, $startAddress, $registerCount) {

        $functionCode = [byte]4
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

 #FC05  Write Single Coil Output
function WriteModbusRTUCoil($slaveAddress, $coilAddress, $value) {
    
    $request = [byte[]]@(
        [byte]$slaveAddress,
        5,
        [System.Convert]::ToByte($coilAddress / 256),
        [System.Convert]::ToByte($coilAddress % 256),
        [System.Convert]::ToByte($value),  # 0 for OFF, 1 for ON
        0  # Reserved byte (usually set to 0)
    )

    
      $request += CalculateCRC $request

    
     $port.Write($request, 0, $request.Length)

    
     Start-Sleep -Milliseconds 100

     
      $response = $port.ReadExisting()
      $responseBytes = [System.Text.Encoding]::ASCII.GetBytes($response)
      if ($responseBytes) {
            Write-Host "Received bytes:"
            $responseBytes -join " " | Write-Host
            return $responseBytes
      } else {
            Write-Host "No response received"
      }

}

#FC06 Write to Single Holding Register
function WriteSingleHoldingRegister($slaveAddress, $registerAddress, $value) {
        $functionCode = [byte]6

        # Construct the request
        $request = [byte[]]@(
            [byte]$slaveAddress,
            $functionCode,
            [System.Convert]::ToByte($registerAddress / 256),
            [System.Convert]::ToByte($registerAddress % 256),
            [System.Convert]::ToByte($value / 256),
            [System.Convert]::ToByte($value % 256)
        )

        # Calculate and append CRC16
        $request += CalculateCRC $request
        $request -join " " | Write-Host

        # Send the request to the Modbus device (you need to set up the port first)
        $port.Write($request, 0, $request.Length)

        # Wait for response (adjust sleep time as needed)
        Start-Sleep -Milliseconds 100

        # Read the response (you need to handle the response accordingly)
        $response = $port.ReadExisting()
        $responseBytes = [System.Text.Encoding]::ASCII.GetBytes($response)

        if ($responseBytes) {
            Write-Host "Received bytes:"
            $responseBytes -join " " | Write-Host
            return $responseBytes
        } else {
            Write-Host "No response received"
        }
}

#FC14  Read Log Record
function ReadSingleFileRecord($slaveAddress, $fileNumber, $startRecord, $recordCount) {
        $functionCode = [byte]20  # 0x14

        # Construct the request
        $request = [byte[]]@(
            [byte]$slaveAddress,
            $functionCode,
            7,
            6,
            [System.Convert]::ToByte($fileNumber / 256),
            [System.Convert]::ToByte($fileNumber % 256),
            [System.Convert]::ToByte($startRecord / 256),
            [System.Convert]::ToByte($startRecord % 256),
            [System.Convert]::ToByte($recordCount / 256),
            [System.Convert]::ToByte($recordCount % 256)
        )

        # Calculate and append CRC16
        $request += CalculateCRC $request
        $request -join " " | Write-Host

        # Send the request to the Modbus device (you need to set up the port first)
        $port.Write($request, 0, $request.Length)

        # Wait for response (adjust sleep time as needed)
        Start-Sleep -Milliseconds 100

        # Read the response (you need to handle the response accordingly)
        $response = $port.ReadExisting()
        $responseBytes = [System.Text.Encoding]::ASCII.GetBytes($response)

        if ($responseBytes) {
            Write-Host "Received bytes:"
            $responseBytes -join " " | Write-Host
            return $responseBytes
        } else {
            Write-Host "No response received"
        }
}

#FC14  Read multi Log Record
function ReadMultiFileRecords($slaveAddress, $fileNumbers, $recordLength) {
    # Calculate the total byte count based on the number of file records
    $byteCount = $fileNumbers.Count * 7

    # Construct the request
    $request = [byte[]]@(
        [byte]$slaveAddress,
        0x14,  # Function code for Read File Record
        $byteCount
    )

    # Add each file record to the request
    foreach ($fileNumber in $fileNumbers) {
        $request += 6  # Reference type (fixed value)
        $request += [System.BitConverter]::GetBytes($fileNumber)  # File number (2 bytes)
        $request += [System.BitConverter]::GetBytes(0)  # Record number (2 bytes, set to 0 for simplicity)
        $request += [System.BitConverter]::GetBytes($recordLength)  # Record length (2 bytes)
    }


     $request += CalculateCRC $request

    
     $port.Write($request, 0, $request.Length)

   
      Start-Sleep -Milliseconds 100

     
      $response = $port.ReadExisting()
      $responseBytes = [System.Text.Encoding]::ASCII.GetBytes($response)
      if ($responseBytes) {
        Write-Host "Received bytes:"
        $responseBytes -join " " | Write-Host
        return $responseBytes
      } else {
            Write-Host "No response received"
      }

}

#FC0F  Read multi Log Record
 function WriteMultipleCoils($slaveAddress, $coilAddresses, $values) {
        # Validate input arrays
        if ($coilAddresses.Length -ne $values.Length) {
            Write-Host "Error: The number of coil addresses must match the number of values."
            return
        }

        # Calculate the total byte count
        $byteCount = [System.Convert]::ToByte($coilAddresses.Length)

        # Construct the request
        $request = [byte[]]@(
            [byte]$slaveAddress,
            0x0F,  # Function code for Write Multiple Coils
            $byteCount
        )

        # Add each coil address and value to the request
        for ($i = 0; $i -lt $coilAddresses.Length; $i++) {
            $request += [System.BitConverter]::GetBytes($coilAddresses[$i])  # Coil address (2 bytes)
            $request += [System.BitConverter]::GetBytes($values[$i])  # Value (2 bytes, 0 for OFF, 1 for ON)
        }

    
        $request += CalculateCRC $request

     
        $port.Write($request, 0, $request.Length)

     
        Start-Sleep -Milliseconds 100

     
        $response = $port.ReadExisting()
        $responseBytes = [System.Text.Encoding]::ASCII.GetBytes($response)


        if ($responseBytes) {
            Write-Host "Received bytes:"
            $responseBytes -join " " | Write-Host
            return $responseBytes
        } else {
            Write-Host "No response received"
        }
}

#  FC10 Write to Multi Holding Register
function WriteMultiRegister ($slaveAddress, $startAddress, $values) {
        $functionCode = [byte]16
        $registerCount = $values.Length   # Assuming each value is 16 bits (2 bytes)
        $byteCount = $values.Length * 2
        [byte[]]$byteValues =$Values

        # Construct the request
        $request = [byte[]]@(
            [byte]$slaveAddress,
            $functionCode,
            [System.Convert]::ToByte($startAddress / 256),
            [System.Convert]::ToByte($startAddress % 256),
            [System.Convert]::ToByte($registerCount / 256),
            [System.Convert]::ToByte($registerCount % 256),
            [System.Convert]::ToByte($byteCount)
            #$byteValues
        )


             # Split each value into high and low bytes
        foreach ($value in $values) {
            $highByte = [System.Convert]::ToByte($value / 256)
            $lowByte = [System.Convert]::ToByte($value % 256)
            $request += $highByte
            $request += $lowByte
        }
            # Calculate and append CRC16
         $request += CalculateCRC $request
         $request -join " " | Write-Host  


        # Send the request to the Modbus device (you need to set up the port first)
        $port.Write($request, 0, $request.Length)

        # Wait for response (adjust sleep time as needed)
        Start-Sleep -Milliseconds 100

        # Read the response (you need to handle the response accordingly)
        $response = $port.ReadExisting()
        $responseBytes = [System.Text.Encoding]::ASCII.GetBytes($response)

        if ($responseBytes) {
            Write-Host "Received bytes:"
            $responseBytes -join " " | Write-Host
            return $responseBytes
        } else {
            Write-Host "No response received"
        }
    }



 

    # Calculate CRC
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

    # Filter bytes for data.
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





















while ($true) {
    if ($Host.UI.RawUI.KeyAvailable -and (3 -eq [int]$Host.UI.RawUI.ReadKey("AllowCtrlC,IncludeKeyUp,NoEcho").Character)) {
        break
        [System.GC]::Collect()  }


    # Serial Port Name
    if ($IsLinux -or [System.Environment]::OSVersion.Platform -eq "Unix"){
        $portName = "/dev/ttyS0"  # replac with real com port 
    } elseif ($IsMacOS -or  [System.Environment]::OSVersion.Platform -eq "Unix") {
        #$portName = "/dev/cu.usbserial-00000000"  # replac with real com port 
        $portName = "/dev/tty.usbserial-00000000"  # replac with real com port 
    } elseif ($IsWindows -or [System.Environment]::OSVersion.Platform -eq "Win32NT") {
        $portName = "COM3"  # replac with real com port 
    }

    $baudRate = 9600  # Replace with the desired baud rate
    $parity = "None"  # Replace with the desired parity (e.g., "None", "Even", "Odd")
    $dataBits = 8  # Replace with the desired number of data bits
    $stopBits = 1  # Replace with the desired number of stop bits (e.g., "One", "OnePointFive", "Two")

    $port = New-Object System.IO.Ports.SerialPort($portName, $baudRate, $parity, $dataBits, $stopBits)

    $port.Open()


             # Write Value example
         
            #WriteHoldingRegister $slaveAddress $startAddress $values
          # $writeBytes =  WriteMultiRegister  2 35840 100,5
           
         
         # Example for Read ABB M1M xx Power Meter

        try {
             # ReadHoldingRegister 2 23296 51
            while ($true) {
                $responseBytes =  ReadHoldingRegister 2 23296 14      # read Voltage
                GetValue $responseBytes   
             
                # write-host $holdingRegisters      
                # Start-Sleep -Seconds 1
                Start-Sleep -Milliseconds 3000 
            }
        } finally {   
            $port.Close()
        }
}










