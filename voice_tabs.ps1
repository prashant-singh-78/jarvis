# Simple voice-to-tab opener using built-in Windows speech engine (no extra Python deps)
# Commands: speak one of the keywords below, e.g. "open youtube"

Add-Type -AssemblyName System.Speech

function Write-Status($msg) {
    $timestamp = (Get-Date -Format "HH:mm:ss")
    Write-Host "[$timestamp] $msg"
}

$commandList = @(
    @{ Keywords = @("whatsapp", "open whatsapp", "whatsapp kholo"); Url = "https://web.whatsapp.com" },
    @{ Keywords = @("instagram", "open instagram", "instagram kholo"); Url = "https://www.instagram.com" },
    @{ Keywords = @("facebook", "open facebook", "facebook kholo"); Url = "https://www.facebook.com" },
    @{ Keywords = @("github", "open github", "git hub", "github kholo"); Url = "https://github.com" },
    @{ Keywords = @("linkedin", "open linkedin", "linked in"); Url = "https://www.linkedin.com" },
    @{ Keywords = @("youtube", "open youtube", "youtube kholo"); Url = "https://www.youtube.com" },
    @{ Keywords = @("spotify", "open spotify", "spotify kholo"); Url = "https://open.spotify.com" },
    @{ Keywords = @("google", "open google"); Url = "https://www.google.com" },
    @{ Keywords = @("exit assistant", "stop listening", "assistant band"); Url = "__exit__" }
)

$installedRecognizers = [System.Speech.Recognition.SpeechRecognitionEngine]::InstalledRecognizers()
$recognizerInfo = $installedRecognizers | Where-Object { $_.Culture.Name -eq "en-US" } | Select-Object -First 1
if (-not $recognizerInfo) { $recognizerInfo = $installedRecognizers | Select-Object -First 1 }
if (-not $recognizerInfo) {
    Write-Status "No speech recognizer is installed. Install Windows speech pack and retry."
    exit 1
}

$recognizer = [System.Speech.Recognition.SpeechRecognitionEngine]::new($recognizerInfo)
$choices = [System.Speech.Recognition.Choices]::new()
$keywordToUrl = @{}

foreach ($cmd in $commandList) {
    foreach ($word in $cmd.Keywords) {
        $choices.Add($word)
        $keywordToUrl[$word.ToLowerInvariant()] = $cmd.Url
    }
}

$grammarBuilder = [System.Speech.Recognition.GrammarBuilder]::new()
$grammarBuilder.Culture = $recognizerInfo.Culture
$grammarBuilder.Append($choices)
$grammar = [System.Speech.Recognition.Grammar]::new($grammarBuilder)
$recognizer.LoadGrammar($grammar)

$recognizer.SetInputToDefaultAudioDevice()

$recognizer.SpeechRecognized += {
    param($sender, $eventArgs)
    $spoken = $eventArgs.Result.Text.ToLowerInvariant()
    if (-not $keywordToUrl.ContainsKey($spoken)) { return }

    $url = $keywordToUrl[$spoken]
    if ($url -eq "__exit__") {
        Write-Status "Stopping assistant per voice command."
        $sender.RecognizeAsyncStop()
        return
    }

    Write-Status "Opening $url"
    Start-Process $url
}

$recognizer.AudioStateChanged += {
    param($sender, $eventArgs)
    Write-Status "Mic state: $($eventArgs.AudioState)"
}

Write-Status "Listening for: $($keywordToUrl.Keys -join ', ')" 
$recognizer.RecognizeAsync([System.Speech.Recognition.RecognizeMode]::Multiple)

try {
    while ($true) { Start-Sleep -Seconds 1 }
}
finally {
    $recognizer.RecognizeAsyncStop()
    $recognizer.Dispose()
    Write-Status "Assistant stopped."
}
