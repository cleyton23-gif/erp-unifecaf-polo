param(
  [Parameter(Mandatory = $true)]
  [string]$VideoPath,

  [string]$OutputDir = "video-frames"
)

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

function Wait-ForUiEvent {
  param(
    [Parameter(Mandatory = $true)]
    [scriptblock]$Register,
    [int]$TimeoutMs = 15000
  )

  $script:waitComplete = $false
  $script:waitTimedOut = $false
  $frame = New-Object System.Windows.Threading.DispatcherFrame
  $timer = New-Object System.Windows.Threading.DispatcherTimer
  $timer.Interval = [TimeSpan]::FromMilliseconds($TimeoutMs)
  $timer.Add_Tick({
    $script:waitTimedOut = $true
    $script:waitComplete = $true
    $timer.Stop()
    $frame.Continue = $false
  })

  $subscription = & $Register {
    $script:waitComplete = $true
    $timer.Stop()
    $frame.Continue = $false
  }

  $timer.Start()
  [System.Windows.Threading.Dispatcher]::PushFrame($frame)

  if ($subscription) {
    Unregister-Event -SubscriptionId $subscription.Id -ErrorAction SilentlyContinue
  }

  if ($script:waitTimedOut) {
    throw "Timed out waiting for media event"
  }
}

function Save-Frame {
  param(
    [System.Windows.Media.MediaPlayer]$Player,
    [int]$Width,
    [int]$Height,
    [string]$Path
  )

  $visual = New-Object System.Windows.Media.DrawingVisual
  $context = $visual.RenderOpen()
  $rect = New-Object System.Windows.Rect(0, 0, $Width, $Height)
  $context.DrawVideo($Player, $rect)
  $context.Close()

  $bitmap = New-Object System.Windows.Media.Imaging.RenderTargetBitmap(
    $Width,
    $Height,
    96,
    96,
    [System.Windows.Media.PixelFormats]::Pbgra32
  )
  $bitmap.Render($visual)

  $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
  $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bitmap))

  $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create)
  try {
    $encoder.Save($stream)
  } finally {
    $stream.Close()
  }
}

$fullVideoPath = [System.IO.Path]::GetFullPath($VideoPath)
$fullOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
[System.IO.Directory]::CreateDirectory($fullOutputDir) | Out-Null

$player = New-Object System.Windows.Media.MediaPlayer
$script:errorText = ""

$player.Add_MediaFailed({
  param($sender, $eventArgs)
  $script:errorText = $eventArgs.ErrorException.Message
})

$player.Open([Uri]$fullVideoPath)

Wait-ForUiEvent -Register {
  param($complete)
  Register-ObjectEvent -InputObject $player -EventName MediaOpened -Action {
    & $Event.MessageData
  } -MessageData $complete
}

if ($script:errorText) {
  throw "Unable to open video: $script:errorText"
}

$duration = $player.NaturalDuration.TimeSpan.TotalSeconds
$width = [Math]::Max(1, $player.NaturalVideoWidth)
$height = [Math]::Max(1, $player.NaturalVideoHeight)

Write-Output (@{
  duration = [Math]::Round($duration, 2)
  width = $width
  height = $height
} | ConvertTo-Json -Compress)

$times = @(
  0.1,
  $duration * 0.12,
  $duration * 0.24,
  $duration * 0.36,
  $duration * 0.48,
  $duration * 0.60,
  $duration * 0.72,
  $duration * 0.84,
  [Math]::Max(0.1, $duration - 0.5)
)

for ($i = 0; $i -lt $times.Count; $i++) {
  $player.Position = [TimeSpan]::FromSeconds($times[$i])
  $player.Play()

  $deadline = [DateTime]::Now.AddMilliseconds(700)
  while ([DateTime]::Now -lt $deadline) {
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
      [System.Windows.Threading.DispatcherPriority]::Background,
      [Action]{}
    )
    Start-Sleep -Milliseconds 20
  }

  $player.Pause()

  $filename = "frame-{0:D2}.png" -f ($i + 1)
  $outputPath = Join-Path $fullOutputDir $filename
  Save-Frame -Player $player -Width $width -Height $height -Path $outputPath
  Write-Output "$filename $([Math]::Round($times[$i], 2))s"
}

$player.Close()
