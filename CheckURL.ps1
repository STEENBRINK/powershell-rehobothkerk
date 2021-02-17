PROCESS {
  if($_) {
	
	$url = $_;

	$urlIsValid = $false
	try
	{
		$request = [System.Net.WebRequest]::Create($url)
		$request.Method = 'HEAD'
		$response = $request.GetResponse()
		$httpStatus = $response.StatusCode
		$urlIsValid = ($httpStatus -eq 'OK')
		$tryError = $null
		$response.Close()
	}
	catch [System.Exception] {
		$httpStatus = $null
		$tryError = $_.Exception
		$urlIsValid = $false;
	}

	$x = new-object Object | `
			add-member -membertype NoteProperty -name IsValid -Value $urlIsvalid -PassThru | `
			add-member -membertype NoteProperty -name Url -Value $_ -PassThru | `
			add-member -membertype NoteProperty -name HttpStatus -Value $httpStatus -PassThru | `
			add-member -membertype NoteProperty -name Error -Value $tryError -PassThru
	$x 
  }
}