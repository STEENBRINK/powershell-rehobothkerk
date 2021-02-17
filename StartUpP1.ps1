$socket = New-Object System.Net.Sockets.TcpClient('', 51234) #add id of companion
$streamdeck_ID = '' #change this first

$stream = $socket.GetStream()
$writer = New-Object System.IO.StreamWriter($stream)

$writer.WriteLine('PAGE-SET 2 ' + $streamdeck_ID)
$writer.Flush()