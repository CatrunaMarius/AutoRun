Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# === DEFINIM APLICAȚIILE PE CATEGORII ===
$appsByCategory = @{
    "Browsere Web"         = @(
        @{ Name = "Google Chrome"; Id = "Google.Chrome" },
        @{ Name = "Mozilla Firefox"; Id = "9NZVDKPMR9RD" },
        @{ Name = "Zen Browser"; Id = "Zen-Team.Zen-Browser" },
        @{ Name = "Brave Browser"; Id = "Brave.Brave" },
        @{Name = "Opera Browser"; Id = "XP8CF6S8G2D5T6" },
        @{Name = "Opera GX"; Id = "XPDBZ4MPRKNN30" },
        @{Name = "Safari"; Id = "Apple.Safari" }
    )
    "Media"                = @(
        @{ Name = "Spotify"; Id = "Spotify.Spotify" },
        @{Name = "SoundCloud"; Id = "9NVJBT29B36L" },
        @{Name = "Deezer Music"; Id = "9NBLGGH6J7VV" },
        @{Name = "iTunes"; Id = "9PB2MZ1ZMB1S" },
        @{Name = "Apple Music"; Id = "9PFHDD62MXS1" },
        @{Name = "Apple TV"; Id = "9NM4T8B9JQZ1" },
        @{Name = "YouTube Music"; Id = "th-ch.YouTubeMusic" },
        @{Name = "Netflix"; Id = "9WZDNCRFJ3TJ" },
        @{Name = "Disney+"; Id = "9NXQXXLFST89" },
        @{Name = "Prime Video for Windows"; Id = "9P6RC76MSMMJ" },
        @{Name = "Amazon Music"; Id = "9WZDNCRFJ3TJ" },
        @{Name = "Screenbox"; Id = "Starpine.Screenbox" },
        @{ Name = "VLC Media Player"; Id = "VideoLAN.VLC" },
        @{ Name = "Winamp"; Id = "Winamp.Winamp " }
    )
    "Games"                = @(
        @{Name = "Epic Games Store"; Id = "XP99VR1BPSBQJ2" },
        @{Name = "Steam"; Id = "9MZTSWTZN7GM" },
        @{Name = "Ubisoft Connect"; Id = "XPDP2QW12DFSFK" },
        @{Name = "Origin"; Id = "ElectronicArts.Origin" }
    )
    "Dev Tools"            = @(
        @{ Name = "Git"; Id = "Git.Git" },
        @{ Name = "NodeJS"; Id = "OpenJS.NodeJS" },
        @{ Name = "Python 3"; Id = "Python.Python.3.13" }, # Python 3
        @{ Name = "Visual Studio Code"; Id = "Microsoft.VisualStudioCode" },
        @{Name = "Visual Studio 2022 Community"; Id = "Microsoft.VisualStudio.2022.Community" },
        @{Name = "Android Studio"; Id = "Google.AndroidStudio" },
        @{Name = "Anaconda3"; Id = "Anaconda.Anaconda3" },
        @{Name = "PyCharm Community Edition"; Id = "JetBrains.PyCharm.Community" },
        @{Name = "Java JDK 24"; Id = "Oracle.JDK.24" },
        @{Name = "Postman"; Id = "Postman.Postman" },
        @{Name = "MongoDB"; Id = "MongoDB.Server" },
        @{Name = "NVIDIA CUDA Toolkit"; Id = "Nvidia.CUDA" },
        @{Name = "Unity 6"; Id = "Unity.Unity.6000" },
        @{ Name = "Docker Desktop"; Id = "XP8CBJ40XLBWKX" },
        @{ Name = "GitHub Desktop"; Id = "GitHub.GitHubDesktop" },
        @{ Name = "GitKraken"; Id = "Axosoft.GitKraken" }
    )
    
    "Suita Office"         = @(
        @{Name = "Adobe Reader"; Id = "Adobe.Acrobat.Reader.64-bit" },
        @{ Name = "Microsoft Office 2021"; Id = "OfficeSetup" }
    )
    "Social"               = @(
        @{Name = "TikTok"; Id = "9NH2GPH4JZS4" },
        @{Name = "Facebook"; Id = "9WZDNCRFJ2WL" },
        @{Name = "Instagram"; Id = "9NBLGGH5L9XT" },
        @{Name = "Threads an Instagram app"; Id = "9MXBP1FB84CQ" },
        @{Name = "Messenger"; Id = "9WZDNCRF0083" },
        @{Name = "WhatsApp"; Id = "9NKSQGP7F2NH" },
        @{Name = "Telegram"; Id = "Telegram.TelegramDesktop" },
        @{Name = "Zoom Workplace"; Id = "XP99J3KP4XZ4VV" },
        @{Name = "Microsoft Teams"; Id = "XP8BT8DW290MPQ" }
    )
    "Arhivare & Utilitare" = @(
        @{Name = "Microsoft PowerToys"; Id = "XP89DCGQ3K6VLD" },
        @{Name = "qBittorrent"; Id = "qBittorrent.qBittorrent" },
        @{ Name = "7-Zip"; Id = "7zip.7zip" },
        @{ Name = "WinRAR"; Id = "RARLab.WinRAR" },
        @{ Name = "ShareX"; Id = "ShareX.ShareX" },
        @{ Name = "Notepad++"; Id = "Notepad++.Notepad++" }
    )
}


$pythonLibraries = @(
    @{ Name = "NumPy"; Id = "Pip_numpy" },
    @{ Name = "Pandas"; Id = "Pip_pandas" },
    @{ Name = "Matplotlib"; Id = "Pip_matplotlib" },
    @{ Name = "Scikit-learn"; Id = "Pip_scikit-learn" },
    @{ Name = "Torch"; Id = "Pip_torch" },
    @{ Name = "OpenCV"; Id = "Pip_opencv-python" },
    @{ Name = "Tensorflow"; Id = "Pip_tensorflow" },
    @{ Name = "Keras"; Id = "Pip_keras" },
    @{ Name = "NLTK"; Id = "Pip_nltk" },
    @{ Name = "Plotly"; Id = "Pip_plotly" },
    @{ Name = "Theano"; Id = "Pip_theano" },
    @{ Name = "Seaborn"; Id = "Pip_seaborn" },
    @{ Name = "Spacy"; Id = "Pip_spacy" },
    @{ Name = "Ultralytics"; Id = "Pip_ultralytics" },
    @{ Name = "Torchvision"; Id = "Pip_torchvision" },
    @{ Name = "Gluoncv"; Id = "Pip_gluoncv" },
    @{ Name = "mmdet"; Id = "Pip_mmdet" },
    @{ Name = "Transformers"; Id = "Pip_transformers" },
    @{ Name = "Diffusers"; Id = "Pip_diffusers" },
    @{ Name = "Jax"; Id = "Pip_jax" },
    @{ Name = "Langchain"; Id = "Pip_langchain" },
    @{ Name = "Llama_index"; Id = "Pip_llama_index" },
    @{ Name = "Fastai"; Id = "Pip_fastai" },
    @{ Name = "Auto-sklearn"; Id = "Pip_auto-sklearn" },
    @{ Name = "Autogluon"; Id = "Pip_autogluon" },
    @{ Name = "Lightgbm"; Id = "Pip_lightgbm" },
    @{ Name = "Xgboost"; Id = "Pip_xgboost" },
    @{ Name = "Pytesseract"; Id = "Pip_pytesseract" },
    @{ Name = "EasyOCR"; Id = "Pip_easyocr" },
    @{ Name = "Keras‑OCR"; Id = "Pip_keras-ocr" },
    @{ Name = "Scikit-image"; Id = "Pip_scikit-image" },
    @{ Name = "Librosa"; Id = "Pip_librosa" },
    @{ Name = "PyAudio"; Id = "Pip_PyAudio" },
    @{ Name = "Speechrecognition"; Id = "Pip_speechrecognition" },
    @{ Name = "Vosk"; Id = "Pip_vosk" },
    @{ Name = "torchaudio"; Id = "Pip_torchaudio" },
    @{ Name = "Openai Whisper"; Id = "Pip_openai-whisper" },
    @{ Name = "Speechbrain"; Id = "Pip_speechbrain" },
    @{ Name = "Pyttsx3"; Id = "Pip_pyttsx3" },
    @{ Name = "Google-cloud-texttospeech"; Id = "Pip_google-cloud-texttospeech" },
    @{ Name = "Coqui-tts"; Id = "Pip_coqui-tts" },
    @{ Name = "Ollama"; Id = "Pip_ollama" },
    @{ Name = "Rasa"; Id = "Pip_rasa" },
    @{ Name = "gradio"; Id = "Pip_gradio" },
    @{ Name = "Streamlit"; Id = "Pip_streamlit" },#
    @{ Name = "ChatterBot"; Id = "Pip_ChatterBot" },
    @{ Name = "Lingua"; Id = "Pip_lingua-language-detector" },
    @{ Name = "Allennlp"; Id = "Pip_allennlp" },
    @{ Name = "Requests"; Id = "Pip_requests" },
    @{ Name = "Beautifulsoup"; Id = "Pip_beautifulsoup4" },
    @{ Name = "Selenium"; Id = "Pip_selenium" },
    @{ Name = "Flask"; Id = "Pip_Flask" },
    @{ Name = "Fastapi"; Id = "Pip_fastapi" },
    @{ Name = "Uvicorn"; Id = "Pip_uvicorn" },
    @{ Name = "Django"; Id = "Pip_Django" },
    @{ Name = "Tornado"; Id = "Pip_tornado" },
    @{ Name = "Sanic"; Id = "Pip_sanic" }

)

# === VARIABILE GLOBALE ===
[int]$script:margin = 12
[int]$script:rowHeight = 18
[int]$script:columnCount = 2 # Numărul de coloane dorit pentru aplicațiile normale
[int]$script:officeComponentColumnCount = 3 # Numărul de coloane dorit pentru componentele Office
[int]$script:pythonLibComponentColumnCount = 3 # Nou: Numărul de coloane dorit pentru bibliotecile Python
[int]$script:columnSpacing = 30 # Spațiu între coloane
[int]$script:checkboxIndent = 20 # Indentare pentru checkbox-uri față de marginea categoriei
[int]$script:officeComponentIndent = 40 # Indentare suplimentară pentru componentele Office
[int]$script:pythonLibComponentIndent = 40 # Nou: Indentare suplimentară pentru bibliotecile Python
[int]$script:pythonEnvOptionIndent = 60 # Nou: Indentare pentru opțiunile mediului Python
[int]$script:categoryTopMargin = 20 # Spațiu de deasupra fiecărei categorii
[int]$script:officeSpaceHeight = 0 # Va stoca înălțimea spațiului ocupat de componentele Office
[int]$script:pythonLibSpaceHeight = 0 # Nou: Va stoca înălțimea spațiului ocupat de bibliotecile Python
[int]$script:pythonEnvOptionHeight = 0 # Nou: Înălțimea spațiului ocupat de opțiunile mediului Python

$script:checkboxes = @()
$script:officeComponentCheckboxes = @()
$script:pythonLibCheckboxes = @() # Nou: Lista pentru checkbox-urile bibliotecilor Python
$script:pythonEnvOptions = @() # Nou: Lista pentru radio buttons Python environment


function New-RoundedRectangleRegion {
    param(
        [int]$x,
        [int]$y,
        [int]$width,
        [int]$height,
        [int]$radius
    )

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    
    # Arcul de sus-stânga
    $path.AddArc($x, $y, $radius * 2, $radius * 2, 180, 90)
    # Linia de sus
    $path.AddLine($x + $radius, $y, $x + $width - $radius, $y)
    
    # Arcul de sus-dreapta
    $path.AddArc($x + $width - $radius * 2, $y, $radius * 2, $radius * 2, 270, 90)
    # Linia din dreapta
    $path.AddLine($x + $width, $y + $radius, $x + $width, $y + $height - $radius)

    # Arcul de jos-dreapta
    $path.AddArc($x + $width - $radius * 2, $y + $height - $radius * 2, $radius * 2, $radius * 2, 0, 90)
    # Linia de jos
    $path.AddLine($x + $width - $radius, $y + $height, $x + $radius, $y + $height)

    # Arcul de jos-stânga
    $path.AddArc($x, $y + $height - $radius * 2, $radius * 2, $radius * 2, 90, 90)
    # Linia din stânga
    $path.AddLine($x, $y + $height - $radius, $x, $y + $radius)

    $path.CloseFigure() # Închide forma

    return (New-Object System.Drawing.Region($path))
}



# === FORMULAR ===
$form = New-Object System.Windows.Forms.Form
$form.Text = "FirstRun App Selector"
$form.Size = New-Object System.Drawing.Size(700, 850)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false

# === TAB CONTROL ===
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$form.Controls.Add($tabControl)

# === TAB 1: Aplicații ===
$tabApps = New-Object System.Windows.Forms.TabPage
$tabApps.Text = "Apps"
$tabControl.TabPages.Add($tabApps)

# Bara de cautare
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Size = New-Object System.Drawing.Size(500, 30)
$searchBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$searchBox.Location = New-Object System.Drawing.Point(80, 30)
$tabApps.Controls.Add($searchBox)

# Panel cu scroll pentru checkbox-uri
$panelApps = New-Object System.Windows.Forms.Panel
$panelApps.Location = New-Object System.Drawing.Point($script:margin, 50)
$panelApps.Size = New-Object System.Drawing.Size(880, 660)
$panelApps.AutoScroll = $true
$tabApps.Controls.Add($panelApps)

# Butoane
$selectAll = New-Object System.Windows.Forms.Button
$selectAll.Text = "Select All"
$selectAll.Size = New-Object System.Drawing.Size(150, 40)
$selectAll.Location = New-Object System.Drawing.Point(30, 730)
$selectAll.ForeColor = [System.Drawing.Color]::DarkGreen
$selectAll.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$selectAll.Region = New-RoundedRectangleRegion 0 0 150 40 20
$tabApps.Controls.Add($selectAll)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "Start Install"
$okButton.Size = New-Object System.Drawing.Size(150, 40)
$okButton.Location = New-Object System.Drawing.Point(250, 730)
$okButton.ForeColor = [System.Drawing.Color]::DarkGreen
$okButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$okButton.Region = New-RoundedRectangleRegion 0 0 150 40 20
$tabApps.Controls.Add($okButton)

$deselectAll = New-Object System.Windows.Forms.Button
$deselectAll.Text = "Deselect All"
$deselectAll.Size = New-Object System.Drawing.Size(150, 40)
$deselectAll.Location = New-Object System.Drawing.Point(480, 730)
$deselectAll.ForeColor = [System.Drawing.Color]::DarkGreen
$deselectAll.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$deselectAll.Region = New-RoundedRectangleRegion 0 0 150 40 20
$tabApps.Controls.Add($deselectAll)

# Funcție pentru a ajusta pozițiile controalelor și AutoScrollMinSize
function Adjust-ControlPositions {
    param(
        [System.Windows.Forms.Panel]$targetPanel,
        [array]$controlsToAdjust,
        [int]$baseY
    )

    [int]$currentY = $baseY
    [int]$margin = [int]$script:margin
    [int]$rowHeight = [int]$script:rowHeight
    [int]$columnCount = [int]$script:columnCount
    [int]$officeComponentColumnCount = [int]$script:officeComponentColumnCount
    [int]$pythonLibComponentColumnCount = [int]$script:pythonLibComponentColumnCount
    [int]$columnSpacing = [int]$script:columnSpacing
    [int]$checkboxIndent = [int]$script:checkboxIndent
    [int]$officeComponentIndent = [int]$script:officeComponentIndent
    [int]$pythonLibComponentIndent = [int]$script:pythonLibComponentIndent
    [int]$pythonEnvOptionIndent = [int]$script:pythonEnvOptionIndent # Nou: Indentare pentru opțiunile mediului Python
    [int]$categoryTopMargin = [int]$script:categoryTopMargin

    # Calculăm lățimea disponibilă pentru coloane și lățimea fiecărei coloane
    [int]$effectivePanelWidth = $targetPanel.Width - (2 * $margin) - $checkboxIndent
    [int]$columnWidth = ($effectivePanelWidth - (($columnCount - 1) * $columnSpacing)) / $columnCount
    if ($columnWidth -lt 1) { $columnWidth = 1 }

    # Lățimea coloanei pentru componentele Office
    [int]$officeEffectivePanelWidth = $targetPanel.Width - (2 * $margin) - $officeComponentIndent
    [int]$officeComponentColumnWidth = ($officeEffectivePanelWidth - (($officeComponentColumnCount - 1) * $columnSpacing)) / $officeComponentColumnCount
    if ($officeComponentColumnWidth -lt 1) { $officeComponentColumnWidth = 1 }

    # Lățimea coloanei pentru bibliotecile Python
    [int]$pythonLibEffectivePanelWidth = $targetPanel.Width - (2 * $margin) - $pythonLibComponentIndent
    [int]$pythonLibComponentColumnWidth = ($pythonLibEffectivePanelWidth - (($pythonLibComponentColumnCount - 1) * $columnSpacing)) / $pythonLibComponentColumnCount
    if ($pythonLibComponentColumnWidth -lt 1) { $pythonLibComponentColumnWidth = 1 }

    [int]$colIndex = 0
    [int]$currentRowYStart = $currentY

    $script:officeSpaceHeight = 0 # Resetăm înălțimea spațiului ocupat de Office components
    $script:pythonLibSpaceHeight = 0 # Resetăm înălțimea spațiului ocupat de bibliotecile Python
    $script:pythonEnvOptionHeight = 0 # Nou: Resetăm înălțimea spațiului ocupat de opțiunile mediului Python

    # Iterăm prin controale pentru a le poziționa
    foreach ($control in $controlsToAdjust) {
        # Dacă este o etichetă de categorie
        if ($control -is [System.Windows.Forms.Label]) {
            $currentY += $categoryTopMargin # Spațiu mare înainte de categoria nouă

            $control.Location = New-Object System.Drawing.Point($margin, $currentY)
            $control.Width = $targetPanel.Width - (2 * $margin)
            $control.AutoSize = $false
            $control.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
            
            $currentY += $rowHeight # Avansăm Y pentru a începe checkbox-urile de sub categorie

            # Resetăm poziția de start pentru checkbox-uri pentru noua categorie
            $colIndex = 0
            $currentRowYStart = $currentY
            continue
        }

        # Gestionăm checkbox-ul principal Office 2021
        if ($control -is [System.Windows.Forms.CheckBox] -and $control.Text -eq "Microsoft Office 2021") {
            $currentX = $margin + $checkboxIndent + ($colIndex % $columnCount) * ($columnWidth + $columnSpacing)

            if ($colIndex % $columnCount -eq 0 -and $colIndex -ne 0) {
                $currentY = $currentRowYStart + $rowHeight
                $currentRowYStart = $currentY
            }
            $control.Location = New-Object System.Drawing.Point($currentX, $currentY)
            
            # Acum, dacă Office 2021 e bifat, poziționăm componentele sale
            if ($control.Checked) {
                [int]$officeCompYStart = $currentY + $rowHeight # Începem sub checkbox-ul Office 2021
                [int]$officeCompColIndex = 0
                [int]$maxOfficeCompY = $officeCompYStart

                foreach ($compChk in $script:officeComponentCheckboxes) {
                    $compChk.Visible = $true
                    $compChk.Enabled = $true

                    # Calculează poziția X pentru coloana de componente Office
                    $compX = $margin + $officeComponentIndent + ($officeCompColIndex % $officeComponentColumnCount) * ($officeComponentColumnWidth + $columnSpacing)
                    
                    # Calculează poziția Y pentru rândul de componente Office
                    [int]$compY = $officeCompYStart + ([System.Math]::Floor($officeCompColIndex / $officeComponentColumnCount) * $rowHeight)

                    $compChk.Location = New-Object System.Drawing.Point($compX, $compY)
                    $officeCompColIndex++
                    $maxOfficeCompY = [System.Math]::Max($maxOfficeCompY, $compY)
                }
                # Calculăm înălțimea spațiului ocupat de componentele Office
                $script:officeSpaceHeight = ($maxOfficeCompY - $currentY) + $rowHeight # Adăugăm un rând pentru spațiul de jos
                if ($script:officeComponentCheckboxes.Count -eq 0) { $script:officeSpaceHeight = 0 } # Caz edge
            }
            else {
                # Ascundem componentele dacă Office 2021 nu e bifat
                $script:officeComponentCheckboxes | ForEach-Object {
                    $_.Visible = $false
                    $_.Enabled = $false
                    $_.Checked = $false # Debifăm dacă sunt ascunse
                }
                $script:officeSpaceHeight = 0 # Nu ocupă spațiu
            }
            
            # Avansăm $currentY, luând în considerare spațiul ocupat de componentele Office
            $currentY += $script:officeSpaceHeight + $rowHeight # Avansăm cu înălțimea spațiului Office + un rând pentru a compensa logica de mai jos
            $currentRowYStart = $currentY # Setăm noua poziție de start a rândului

            $colIndex = 0 # Resetăm indicele coloanei pentru următoarele aplicații după blocul Office
            continue # Trecem la următorul control (nu un checkbox Office Component)
        }

        # Nou: Gestionăm checkbox-ul principal Python 3
        if ($control -is [System.Windows.Forms.CheckBox] -and $control.Text -eq "Python 3") {
            $currentX = $margin + $checkboxIndent + ($colIndex % $columnCount) * ($columnWidth + $columnSpacing)

            if ($colIndex % $columnCount -eq 0 -and $colIndex -ne 0) {
                $currentY = $currentRowYStart + $rowHeight
                $currentRowYStart = $currentY
            }
            $control.Location = New-Object System.Drawing.Point($currentX, $currentY)

            # Acum, dacă Python 3 e bifat, poziționăm bibliotecile și opțiunile de mediu
            if ($control.Checked) {
                [int]$pythonLibYStart = $currentY + $rowHeight # Începem sub checkbox-ul Python 3
                [int]$pythonLibColIndex = 0
                [int]$maxPythonLibY = $pythonLibYStart

                foreach ($libChk in $script:pythonLibCheckboxes) {
                    $libChk.Visible = $true
                    $libChk.Enabled = $true

                    # Calculează poziția X pentru coloana de biblioteci Python
                    $libX = $margin + $pythonLibComponentIndent + ($pythonLibColIndex % $pythonLibComponentColumnCount) * ($pythonLibComponentColumnWidth + $columnSpacing)
                    
                    # Calculează poziția Y pentru rândul de biblioteci Python
                    [int]$libY = $pythonLibYStart + ([System.Math]::Floor($pythonLibColIndex / $pythonLibComponentColumnCount) * $rowHeight)

                    $libChk.Location = New-Object System.Drawing.Point($libX, $libY)
                    $pythonLibColIndex++
                    $maxPythonLibY = [System.Math]::Max($maxPythonLibY, $libY)
                }
                # Calculăm înălțimea spațiului ocupat de bibliotecile Python
                $script:pythonLibSpaceHeight = ($maxPythonLibY - $currentY) + $rowHeight
                if ($script:pythonLibCheckboxes.Count -eq 0) { $script:pythonLibSpaceHeight = 0 }
                
                # NOU: Gestionăm radio button-urile pentru mediul Python
                [int]$pythonEnvYStart = $currentY + $script:pythonLibSpaceHeight + $rowHeight # Începem sub spațiul bibliotecilor Python
                [int]$pythonEnvColIndex = 0
                [int]$maxPythonEnvY = $pythonEnvYStart

                foreach ($envRb in $script:pythonEnvOptions) {
                    $envRb.Visible = $true
                    $envRb.Enabled = $true

                    # Calculează poziția X pentru coloana de opțiuni de mediu Python
                    $envX = $margin + $pythonEnvOptionIndent + ($pythonEnvColIndex % $script:columnCount) * ($columnWidth + $columnSpacing) # Folosim columnCount general pentru opțiuni
                    
                    # Calculează poziția Y pentru rândul de opțiuni de mediu Python
                    [int]$envY = $pythonEnvYStart + ([System.Math]::Floor($pythonEnvColIndex / $script:columnCount) * $rowHeight)

                    $envRb.Location = New-Object System.Drawing.Point($envX, $envY)
                    $pythonEnvColIndex++
                    $maxPythonEnvY = [System.Math]::Max($maxPythonEnvY, $envY)
                }
                # Calculăm înălțimea spațiului ocupat de opțiunile de mediu Python
                $script:pythonEnvOptionHeight = ($maxPythonEnvY - $currentY) + $rowHeight
                if ($script:pythonEnvOptions.Count -eq 0) { $script:pythonEnvOptionHeight = 0 }

            }
            else {
                # Ascundem bibliotecile și opțiunile de mediu dacă Python 3 nu e bifat
                $script:pythonLibCheckboxes | ForEach-Object {
                    $_.Visible = $false
                    $_.Enabled = $false
                    $_.Checked = $false # Debifăm dacă sunt ascunse
                }
                $script:pythonLibSpaceHeight = 0

                $script:pythonEnvOptions | ForEach-Object {
                    $_.Visible = $false
                    $_.Enabled = $false
                    $_.Checked = $false
                }
                $script:pythonEnvOptionHeight = 0
            }

            # Avansăm $currentY, luând în considerare spațiul ocupat de biblioteci și opțiunile de mediu Python
            $currentY += $script:pythonLibSpaceHeight + $script:pythonEnvOptionHeight + $rowHeight
            $currentRowYStart = $currentY

            $colIndex = 0 # Resetăm indicele coloanei
            continue # Trecem la următorul control
        }

        # Pentru toate celelalte checkbox-uri (non-categorie, non-Office principal, non-Python principal, non-Office component, non-Python component, non-PythonEnv option)
        # Sări peste componentele Office, Python Libraries și opțiunile de mediu Python, ele sunt gestionate în logica de mai sus
        if ($control.Tag -is [string] -and ($control.Tag.StartsWith("OfficeComponent_") -or $control.Tag.StartsWith("Pip_") -or $control.Tag.StartsWith("PythonEnv_"))) {
            continue
        }

        # Calculează poziția X pentru coloana curentă
        $currentX = $margin + $checkboxIndent + ($colIndex % $columnCount) * ($columnWidth + $columnSpacing)

        # Dacă suntem la începutul unui nou rând (dar nu prima dată deloc)
        if ($colIndex % $columnCount -eq 0 -and $colIndex -ne 0) {
            $currentY = $currentRowYStart + $rowHeight # Avansăm Y pentru rândul următor
        }

        $control.Location = New-Object System.Drawing.Point($currentX, $currentY)
        $colIndex++ # Trecem la următoarea coloană

        if ($colIndex % $columnCount -eq 0) {
            # Dacă am umplut rândul curent
            $currentY += $rowHeight # Trecem la rândul următor
            $currentRowYStart = $currentY # Setăm noua poziție de start a rândului
        }
    }

    # Asigurăm că AutoScrollMinSize reflectă înălțimea maximă atinsă
    [int]$finalY = $currentY
    if ($colIndex % $columnCount -ne 0) {
        $finalY = $currentRowYStart + $rowHeight
    }
    # Dacă ultimul element a fost Office 2021 sau Python 3, asigură-te că înălțimea include componentele sale
    if ($script:officeSpaceHeight -gt 0 -or $script:pythonLibSpaceHeight -gt 0 -or $script:pythonEnvOptionHeight -gt 0) {
        $finalY = [System.Math]::Max($finalY, $currentY)
    }

    $targetPanel.AutoScrollMinSize = New-Object System.Drawing.Size(0, ([int]$finalY + [int]$margin))
    $targetPanel.Invalidate() # Forțează re-draw-ul
}


# Funcție de refresh listă aplicații
function Refresh-AppList([string]$filter = "") {
    $panelApps.Controls.Clear()
    $localCheckboxes = @()
    $script:officeComponentCheckboxes = @()
    $script:pythonLibCheckboxes = @() # Nou

    # Listă temporară pentru a stoca controalele în ordinea dorită, înainte de a le adăuga în panel
    $orderedControls = @()

    foreach ($category in $appsByCategory.Keys) {
        $categoryApps = $appsByCategory[$category] | Where-Object { $_.Name -like "*$filter*" }
        if ($categoryApps.Count -eq 0) { continue }

        $catLabel = New-Object System.Windows.Forms.Label
        $catLabel.Text = $category
        $catLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
        $catLabel.ForeColor = [System.Drawing.Color]::Navy
        $catLabel.AutoSize = $true
        $orderedControls += $catLabel

        foreach ($app in $categoryApps) {
            $cb = New-Object System.Windows.Forms.CheckBox
            $cb.Text = $app.Name
            $cb.Tag = $app.Id
            $cb.AutoSize = $true
            $cb.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            $orderedControls += $cb
            $localCheckboxes += $cb

            if ($app.Name -eq "Microsoft Office 2021") {
                $cb.Add_CheckedChanged({
                        Adjust-ControlPositions -targetPanel $panelApps -controlsToAdjust $panelApps.Controls -baseY 10
                    })

                $officeComponents = @("Word", "Excel", "PowerPoint", "Outlook", "Access", "Publisher")
                foreach ($comp in $officeComponents) {
                    $compChk = New-Object System.Windows.Forms.CheckBox
                    $compChk.Text = $comp
                    $compChk.Tag = "OfficeComponent_$comp"
                    $compChk.AutoSize = $true
                    $compChk.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                    $compChk.Visible = $false
                    $compChk.Enabled = $false
                    $orderedControls += $compChk
                    $script:officeComponentCheckboxes += $compChk
                    $localCheckboxes += $compChk
                }
            }
            # Nou: Adaugă evenimentul CheckedChanged pentru Python 3

            elseif ($app.Name -eq "Python 3") {
                $cb.Add_CheckedChanged({
                        # Când se schimbă starea Python 3, ajustăm pozițiile și vizibilitatea opțiunilor de mediu
                        if ($_.Checked) {
                            # Afișează opțiunile de mediu și le bifează pe cea globală implicit
                            $script:pythonEnvOptions | ForEach-Object { $_.Visible = $true; $_.Enabled = $true }
                            $script:pythonEnvOptions | Where-Object { $_.Tag -eq "PythonEnv_Global" } | ForEach-Object { $_.Checked = $true }
                        }
                        else {
                            # Ascunde și debifează opțiunile de mediu și bibliotecile
                            $script:pythonEnvOptions | ForEach-Object { $_.Visible = $false; $_.Enabled = $false; $_.Checked = $false }
                            $script:pythonLibCheckboxes | ForEach-Object { $_.Visible = $false; $_.Enabled = $false; $_.Checked = $false }
                        }
                        Adjust-ControlPositions -targetPanel $panelApps -controlsToAdjust $panelApps.Controls -baseY 10
                    })

                # --- NOU: Adaugă Radio Buttons pentru opțiuni de mediu Python ---
                $rbGlobal = New-Object System.Windows.Forms.RadioButton
                $rbGlobal.Text = "Install libraries globally (not recommended)"
                $rbGlobal.Tag = "PythonEnv_Global" # Tag pentru identificare
                $rbGlobal.AutoSize = $true
                $rbGlobal.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                $rbGlobal.Visible = $false # Inițial invizibil
                $rbGlobal.Enabled = $false # Inițial inactiv
                $orderedControls += $rbGlobal
                $script:pythonEnvOptions += $rbGlobal

                $rbVirtual = New-Object System.Windows.Forms.RadioButton
                $rbVirtual.Text = "Create new virtual environment"
                $rbVirtual.Tag = "PythonEnv_Virtual" # Tag pentru identificare
                $rbVirtual.AutoSize = $true
                $rbVirtual.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                $rbVirtual.Visible = $false # Inițial invizibil
                $rbVirtual.Enabled = $false # Inițial inactiv
                $orderedControls += $rbVirtual
                $script:pythonEnvOptions += $rbVirtual

                # Adaugă un event handler pentru a actualiza pozițiile când se schimbă selecția mediului
                $rbGlobal.Add_CheckedChanged({
                        if ($rbGlobal.Checked) { Adjust-ControlPositions -targetPanel $panelApps -controlsToAdjust $panelApps.Controls -baseY 10 }
                    })
                $rbVirtual.Add_CheckedChanged({
                        if ($rbVirtual.Checked) { Adjust-ControlPositions -targetPanel $panelApps -controlsToAdjust $panelApps.Controls -baseY 10 }
                    })
            
                # Adaugă checkbox-urile pentru bibliotecile Python
                foreach ($lib in $pythonLibraries) {
                    $libChk = New-Object System.Windows.Forms.CheckBox
                    $libChk.Text = $lib.Name
                    $libChk.Tag = $lib.Id
                    $libChk.AutoSize = $true
                    $libChk.Font = New-Object System.Drawing.Font("Segoe UI", 9)
                    $libChk.Visible = $false # Inițial invizibile
                    $libChk.Enabled = $false  # Inițial inactive
                    $orderedControls += $libChk # Adaugă la lista ordonată
                    $script:pythonLibCheckboxes += $libChk # Adăugăm la lista globală
                    $localCheckboxes += $libChk # Adăugăm și la lista locală pentru Select All/Deselect All
                }
            }
           
        }
    }

    # Adăugăm toate controalele în panel și apoi ajustăm pozițiile inițiale
    foreach ($control in $orderedControls) {
        $panelApps.Controls.Add($control)
    }

    $script:checkboxes = $localCheckboxes

    # Apel inițial pentru a seta pozițiile și vizibilitatea
    Adjust-ControlPositions -targetPanel $panelApps -controlsToAdjust $orderedControls -baseY 10
}

# Evenimente butoane
$searchBox.Add_TextChanged({ Refresh-AppList $searchBox.Text })
$selectAll.Add_Click({
        $script:checkboxes | ForEach-Object {
            if ($_.Visible) {
                $_.Checked = $true
            }
        }
        Adjust-ControlPositions -targetPanel $panelApps -controlsToAdjust $panelApps.Controls -baseY 10
    })
$deselectAll.Add_Click({
        $script:checkboxes | ForEach-Object { $_.Checked = $false }
        Adjust-ControlPositions -targetPanel $panelApps -controlsToAdjust $panelApps.Controls -baseY 10
    })





function Generate-OfficeXml {
    param([array]$SelectedComponents)

    Write-Host "Generare XML pentru Office cu componente: $($SelectedComponents -join ', ')"
    $officeVersion = "ProPlus2021Volume"
    $channel = "PerpetualVL2021"
    $language = "en-us"
    $xmlFilePath = Join-Path -Path $PSScriptRoot -ChildPath "OfficeConfig.xml"

    $allApps = @("Word", "Excel", "PowerPoint", "Outlook", "Publisher", "Access", "OneNote")

    $xml = New-Object System.Xml.XmlDocument

    $configuration = $xml.CreateElement("Configuration")
    $xml.AppendChild($configuration) | Out-Null

    $add = $xml.CreateElement("Add")
    $add.SetAttribute("OfficeClientEdition", "64")
    $add.SetAttribute("Channel", $channel)
    $configuration.AppendChild($add) | Out-Null

    $product = $xml.CreateElement("Product")
    $product.SetAttribute("ID", $officeVersion)
    $add.AppendChild($product) | Out-Null

    $languageNode = $xml.CreateElement("Language")
    $languageNode.SetAttribute("ID", $language)
    $product.AppendChild($languageNode) | Out-Null

    foreach ($app in $allApps) {
        if (-not ($SelectedComponents -contains $app)) {
            $excludeApp = $xml.CreateElement("ExcludeApp")
            $excludeApp.SetAttribute("ID", $app)
            $product.AppendChild($excludeApp) | Out-Null
        }
    }

    # Properties
    $props = @(
        @{ Name = "SharedComputerLicensing"; Value = "0" },
        @{ Name = "FORCEAPPSHUTDOWN"; Value = "FALSE" },
        @{ Name = "AUTOACTIVATE"; Value = "0" },
        @{ Name = "DeviceBasedLicensing"; Value = "0" }
    )

    foreach ($prop in $props) {
        $property = $xml.CreateElement("Property")
        $property.SetAttribute("Name", $prop.Name)
        $property.SetAttribute("Value", $prop.Value)
        $configuration.AppendChild($property) | Out-Null
    }

    # Updates
    $updates = $xml.CreateElement("Updates")
    $updates.SetAttribute("Enabled", "TRUE")
    $updates.SetAttribute("Channel", $channel)
    $configuration.AppendChild($updates) | Out-Null

    # Display
    $display = $xml.CreateElement("Display")
    $display.SetAttribute("Level", "None")
    $display.SetAttribute("AcceptEULA", "TRUE")
    $configuration.AppendChild($display) | Out-Null

    try {
        $xml.Save($xmlFilePath)
        Write-Host "Fisierul OfficeConfig.xml a fost salvat la $xmlFilePath"
        return $xmlFilePath
    }
    catch {
        Write-Error "Eroare la salvarea XML: $($_.Exception.Message)"
        return $null
    }
}


function Download-OfficeSetupExe {
    $url = 'https://github.com/CatrunaMarius/autoInstall/raw/refs/heads/main/bin.exe'
    $path = Join-Path -Path $PSScriptRoot -ChildPath 'bin.exe'
    Invoke-WebRequest -Uri $url -OutFile $path -UseBasicParsing
       
    return $path #Returnează calea reală a fișierului executabil
}


$okButton.Add_Click({
        $selectedApps = $script:checkboxes | Where-Object { $_.Checked -and $_.Tag -notlike "OfficeComponent_*" -and $_.Tag -notlike "Pip_*" } | ForEach-Object { $_.Text }

        # Verificăm dacă Office 2021 a fost selectat (doar checkbox-ul principal)
        $office2021Checkbox = $script:checkboxes | Where-Object { $_.Text -eq "Microsoft Office 2021" } | Select-Object -First 1
        $office2021Selected = $false
        if ($office2021Checkbox -and $office2021Checkbox.Checked) {
            $office2021Selected = $true
        }

        # Verificăm dacă Python 3 a fost selectat
        $python3Checkbox = $script:checkboxes | Where-Object { $_.Text -eq "Python 3" } | Select-Object -First 1
        $python3Selected = $false
        if ($python3Checkbox -and $python3Checkbox.Checked) {
            $python3Selected = $true
        }

        if ($selectedApps.Count -eq 0 -and -not $office2021Selected -and -not $python3Selected) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one app to install.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
   
        # === Inițializare bară de progres și liste status ===
        $totalSteps = $selectedApps.Count
        if ($office2021Selected) { $totalSteps += 1 } # Un pas pentru Office
        if ($python3Selected) { $totalSteps += 1 }    # Un pas pentru Python de bază + un pas pentru fiecare bibliotecă Pip
    
        # Pentru bibliotecile Python, trebuie să numărăm câte sunt
        if ($python3Selected) {
            $selectedPythonLibrariesCount = ($script:pythonLibCheckboxes | Where-Object { $_.Checked -and $_.Visible -and $_.Enabled }).Count
            $totalSteps += $selectedPythonLibrariesCount # Adăugăm un pas pentru fiecare bibliotecă Python
        }


        $formProgress = New-Object System.Windows.Forms.Form
        $formProgress.Text = "Instalare aplicații..."
        $formProgress.Size = New-Object System.Drawing.Size(400, 100)
        $formProgress.StartPosition = "CenterScreen"
        $formProgress.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog # Împiedică redimensionarea
        $formProgress.ControlBox = $false # Elimină butoanele de închidere/minimizare

        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Minimum = 0
        $progressBar.Maximum = $totalSteps
        $progressBar.Step = 1
        $progressBar.Style = 'Continuous'
        $progressBar.Width = 360
        $progressBar.Height = 20
        $progressBar.Location = New-Object System.Drawing.Point(10, 20)

        $labelProgress = New-Object System.Windows.Forms.Label
        $labelProgress.Text = "Pregătire instalare..."
        $labelProgress.Location = New-Object System.Drawing.Point(10, 45)
        $labelProgress.AutoSize = $true
        $labelProgress.Width = 360 # Setează o lățime fixă pentru a evita problemele de layout

        $formProgress.Controls.Add($progressBar)
        $formProgress.Controls.Add($labelProgress)
        $formProgress.Topmost = $true
        $formProgress.Show()
        $formProgress.Refresh() # Forțează redesenarea ferestrei

        # === Liste pentru aplicații instalate și cele care au eșuat ===
        $installedApps = @()
        $failedApps = @()

       


        # === LOGICĂ SPECIALĂ PENTRU OFFICE 2021 ===
        if ($office2021Selected) {
            $labelProgress.Text = "Instalare Microsoft Office 2021..."
            $formProgress.Refresh()

            $selectedOfficeComponents = $script:officeComponentCheckboxes | Where-Object { $_.Checked -and $_.Visible -and $_.Enabled } | ForEach-Object { $_.Text }

            if ($selectedOfficeComponents.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Ați selectat Office 2021, dar nu ați ales componentele. Selectați cel puțin una.", "Avertisment", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $progressBar.PerformStep() # Marchează un pas chiar dacă a eșuat sau a fost anulat
                $failedApps += "Microsoft Office 2021 (componente nespecificate)"
                return
            }

            # Generare și descărcare (presupunem că aceste funcții returnează căi valide)
            $xmlPath = Generate-OfficeXml -SelectedComponents $selectedOfficeComponents
            $setupExePath = Download-OfficeSetupExe

            if ((Test-Path $setupExePath) -and (Test-Path $xmlPath)) {
                # Scoateți MessageBox-ul de aici
                # [System.Windows.Forms.MessageBox]::Show("Instalare Microsoft Office 2021 a început. Vă rugăm așteptați.", "Instalare Office", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
                Write-Host "DEBUG: Pornind instalarea Office..."
                $officeProcess = Start-Process -FilePath $setupExePath -ArgumentList "/configure `"$xmlPath`"" -Wait -PassThru -NoNewWindow
            
                if ($officeProcess.ExitCode -eq 0) {
                    # Scoateți MessageBox-ul de aici
                    # [System.Windows.Forms.MessageBox]::Show("Instalarea Microsoft Office 2021 a fost finalizată cu succes!", "Instalare Office", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    $installedApps += "Microsoft Office 2021"
                    Write-Host "INFO: Microsoft Office 2021 instalat cu succes."
                }
                else {
                    # Scoateți MessageBox-ul de aici
                    # [System.Windows.Forms.MessageBox]::Show("Eroare la instalarea Microsoft Office 2021. Cod de ieșire: $($officeProcess.ExitCode). Verifică manual.", "Eroare Instalare Office", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    $failedApps += "Microsoft Office 2021 (Cod de ieșire: $($officeProcess.ExitCode))"
                    Write-Error "ERROR: Eroare la instalarea Microsoft Office 2021. Cod de ieșire: $($officeProcess.ExitCode)."
                }
            }
            else {
                # Scoateți MessageBox-ul de aici
                # [System.Windows.Forms.MessageBox]::Show("Fișierele de instalare Office nu au fost găsite. Asigură-te că sunt descărcate corect.", "Eroare Fișiere Office", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $failedApps += "Microsoft Office 2021 (Fișiere instalare negăsite)"
                Write-Error "ERROR: Fișierele de instalare Office nu au fost găsite."
            }
            $progressBar.PerformStep()
            $formProgress.Refresh()
        }

        # --- NOU: LOGICĂ SPECIALĂ PENTRU BIBLIOTECILE PYTHON ---
        if ($python3Selected) {
            $labelProgress.Text = "Instalare Python 3 și biblioteci..."
            $formProgress.Refresh()

            $selectedPythonLibraries = $script:pythonLibCheckboxes | Where-Object { $_.Checked -and $_.Visible -and $_.Enabled } | ForEach-Object { $_.Tag.Replace("Pip_", "") }
            $selectedPythonEnvOption = $script:pythonEnvOptions | Where-Object { $_.Checked } | Select-Object -First 1

            if (-not $selectedPythonEnvOption) {
                [System.Windows.Forms.MessageBox]::Show("Ați selectat Python 3, dar nu ați ales o opțiune de mediu pentru biblioteci (Global/Virtual Environment).", "Avertisment", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                $progressBar.PerformStep() # Marchează un pas chiar dacă a eșuat sau a fost anulat
                $failedApps += "Python 3 (opțiune mediu nespecificată)"
                return
            }
       
            # Implementarea reală: Instalare Python de bază (dacă nu e deja instalat)
            try {
                $pythonPath = $null 
                Write-Host "DEBUG: Initializing `$pythonPath` to: '$pythonPath'"

                $cmdResult = Get-Command python.exe -ErrorAction SilentlyContinue
                if ($cmdResult -and $cmdResult.Source -and (Test-Path $cmdResult.Source) -and ($cmdResult.Source -notlike "*Microsoft\WindowsApps*")) {
                    $pythonPath = $cmdResult.Source
                    Write-Host "DEBUG: Python 3 found in PATH via Get-Command: '$pythonPath'"
                }

                if (-not $pythonPath) {
                    $tempWingetOutput = [System.IO.Path]::GetTempFileName()
                    $tempWingetError = [System.IO.Path]::GetTempFileName()

                    # Scoateți MessageBox-ul de aici
                    # [System.Windows.Forms.MessageBox]::Show("Instalare Python 3 prin Winget a început. Vă rugăm așteptați...", "Instalare Python", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                
                    Write-Host "DEBUG: Installing Python 3 via Winget..."
                    $wingetProcess = Start-Process -FilePath "winget" -ArgumentList "install Python.Python.3.13 --silent --accept-source-agreements --accept-package-agreements" -Wait -NoNewWindow -PassThru -RedirectStandardOutput $tempWingetOutput -RedirectStandardError $tempWingetError

                    if ($wingetProcess.ExitCode -ne 0) {
                        $wingetErrorContent = Get-Content $tempWingetError | Out-String 
                        Write-Error "Winget Python installation failed. Error: $wingetErrorContent"
                        # Scoateți MessageBox-ul de aici
                        # [System.Windows.Forms.MessageBox]::Show("Eroare la instalarea Python 3 prin Winget. Vezi consola PowerShell pentru detalii." + "`n" + $wingetErrorContent, "Eroare Winget", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        $failedApps += "Python 3 (Winget eșec)"
                        Remove-Item $tempWingetOutput, $tempWingetError -ErrorAction SilentlyContinue
                        $progressBar.PerformStep() # Marchează un pas chiar dacă a eșuat
                        $formProgress.Refresh()
                        return
                    }
                    $wingetOutputContent = Get-Content $tempWingetOutput | Out-String 
                    Write-Host "DEBUG: Winget Python Output: $wingetOutputContent"
                    Remove-Item $tempWingetOutput, $tempWingetError -ErrorAction SilentlyContinue
                    Write-Host "DEBUG: Python 3 Winget installation finished."
               
                    Start-Sleep -Seconds 5 
                    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + "; " + [System.Environment]::GetEnvironmentVariable("Path", "User")
                    Write-Host "DEBUG: PATH refreshed after Winget installation."

                    $cmdResult = Get-Command python.exe -ErrorAction SilentlyContinue
                    if ($cmdResult -and $cmdResult.Source -and (Test-Path $cmdResult.Source) -and ($cmdResult.Source -notlike "*Microsoft\WindowsApps*")) {
                        $pythonPath = $cmdResult.Source
                        Write-Host "DEBUG: Python 3 found in PATH after re-check: '$pythonPath'"
                    }
                }

                if (-not $pythonPath) {
                    Write-Host "DEBUG: Python not found in PATH. Searching common install locations..."
                    $commonPythonBaseDirs = @(
                        Join-Path $env:LOCALAPDATA "Programs\Python" 
                        "C:\Program Files" 
                        "C:\Python" 
                    )

                    foreach ($baseDir in $commonPythonBaseDirs) {
                        $foundPythonExe = Get-ChildItem -Path $baseDir -Filter "python.exe" -Recurse -ErrorAction SilentlyContinue | 
                        Where-Object { 
                            $_.FullName -like "*Python*\python.exe" -and 
                            $_.DirectoryName -notlike "*Microsoft\WindowsApps*" 
                        } | 
                        Sort-Object LastWriteTime -Descending | Select-Object -First 1
                        if ($foundPythonExe) {
                            $pythonPath = $foundPythonExe.FullName
                            Write-Host "DEBUG: Found Python executable at fallback path: '$pythonPath'"
                            break 
                        }
                    }
                }
                else {
                    Write-Host "DEBUG: Python 3 is already installed at: '$pythonPath'"
                }

                $pythonExecutable = $pythonPath 
           
                if ([string]::IsNullOrEmpty($pythonExecutable) -or (-not (Test-Path $pythonExecutable))) {
                    Write-Error "ERROR: Python executable not found or invalid: '$pythonExecutable'"
                    # Scoateți MessageBox-ul de aici
                    # [System.Windows.Forms.MessageBox]::Show("Nu s-a putut găsi executabilul Python după instalare sau căutare. Asigură-te că alias-urile 'python.exe' sunt dezactivate în Setările Windows (App execution aliases).", "Python Nu a fost Găsit", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    $failedApps += "Python 3 (executabil negăsit)"
                    $progressBar.PerformStep() # Marchează un pas chiar dacă a eșuat
                    $formProgress.Refresh()
                    return 
                }
                Write-Host "DEBUG: Python executable confirmed for operations: '$pythonExecutable'"
                # Scoateți MessageBox-ul de aici
                # [System.Windows.Forms.MessageBox]::Show("Python 3 a fost detectat/instalat cu succes la: `n$pythonExecutable`nSe continuă cu instalarea bibliotecilor...", "Python Găsit/Instalat", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $installedApps += "Python 3" # Considerăm Python de bază instalat aici
                $progressBar.PerformStep() # Pas pentru instalarea Python de bază
                $formProgress.Refresh()

                # NOU: Adaugă/Corectează calea Python și Scripts în PATH la nivel de utilizator
                $pythonDir = Split-Path $pythonExecutable
                $pythonScriptsDir = Join-Path $pythonDir "Scripts"
                Write-Host "DEBUG: Python directory: '$pythonDir', Python Scripts directory: '$pythonScriptsDir'"

                $currentPath = [System.Environment]::GetEnvironmentVariable("Path", "User") 
                $pathElements = $currentPath -split ';' | Where-Object { -not [string]::IsNullOrEmpty($_) }

                $pathChanged = $false
                if (-not ($pathElements -contains $pythonDir)) {
                    $pathElements += $pythonDir
                    $pathChanged = $true
                    Write-Host "DEBUG: Adding Python install directory to User PATH: '$pythonDir'"
                }
                if (-not ($pathElements -contains $pythonScriptsDir)) {
                    $pathElements += $pythonScriptsDir
                    $pathChanged = $true
                    Write-Host "DEBUG: Adding Python Scripts directory to User PATH: '$pythonScriptsDir'"
                }

                if ($pathChanged) {
                    $newPath = ($pathElements | Select-Object -Unique) -join ';'
                    [System.Environment]::SetEnvironmentVariable("Path", $newPath, "User")
                    Write-Host "DEBUG: User PATH updated successfully. New PATH: '$newPath'"
                    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + "; " + [System.Environment]::GetEnvironmentVariable("Path", "User")
                    Start-Sleep -Seconds 2 
                }
                else {
                    Write-Host "DEBUG: Python directories are already in User PATH."
                }

            }
            catch {
                Write-Error "ERROR in Python installation/detection block: $_"
                # Scoateți MessageBox-ul de aici
                # [System.Windows.Forms.MessageBox]::Show("A apărut o eroare la instalarea sau localizarea Python 3. Error: $_", "Python Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                $failedApps += "Python 3 (Eroare internă: $($_.Exception.Message))"
                $progressBar.PerformStep() # Marchează un pas chiar dacă a eșuat
                $formProgress.Refresh()
                return
            }

            # Acum, continuăm cu instalarea bibliotecilor, știind că $pythonExecutable este valid
            # și căile Python/Pip sunt în PATH-ul utilizatorului (și al sesiunii curente)

            if ($selectedPythonLibraries.Count -gt 0) {
                $envType = $selectedPythonEnvOption.Tag 
                Write-Host "DEBUG: Selected Python environment type for libraries: '$envType'"

                if ($envType -eq "PythonEnv_Global") {
                    # Scoateți MessageBox-ul de aici
                    # [System.Windows.Forms.MessageBox]::Show("Instalare biblioteci Python Global: $($selectedPythonLibraries -join ', ')", "Instalare Python Global", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
               
                    try {
                        $globalPipPath = (Get-Command pip.exe -ErrorAction SilentlyContinue).Source
                        Write-Host "DEBUG: Global Pip path (from Get-Command): '$globalPipPath'"
                        if (-not $globalPipPath) {
                            $globalPipPath = Join-Path (Split-Path $pythonExecutable) "Scripts\pip.exe"
                            Write-Host "DEBUG: Fallback Global Pip path: '$globalPipPath'"
                        }

                        if (Test-Path $globalPipPath) {
                            Write-Host "DEBUG: Global Pip executable found at: '$globalPipPath'"
                            foreach ($lib in $selectedPythonLibraries) {
                                $labelProgress.Text = "Instalare bibliotecă Python (global): $lib..."
                                $formProgress.Refresh()

                                Write-Host "DEBUG: Installing '$lib' globally using '$globalPipPath'..."
                                $tempOutput = [System.IO.Path]::GetTempFileName()
                                $tempError = [System.IO.Path]::GetTempFileName()

                                $process = Start-Process -FilePath $globalPipPath -ArgumentList "install $lib" -NoNewWindow -Wait -PassThru -RedirectStandardOutput $tempOutput -RedirectStandardError $tempError

                                if ($process.ExitCode -ne 0) {
                                    $errorOutput = Get-Content $tempError | Out-String 
                                    Write-Error "Failed to install '$lib' globally. Error: $errorOutput"
                                    # Scoateți MessageBox-ul de aici
                                    # [System.Windows.Forms.MessageBox]::Show("Eroare la instalarea '$lib' global.`nDetalii: " + $errorOutput, "Eroare Instalare", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                                    $failedApps += "Python Library (Global): $lib"
                                }
                                else {
                                    $outputContent = Get-Content $tempOutput | Out-String 
                                    Write-Host "Successfully installed '$lib' globally. Output: $outputContent"
                                    # Scoateți MessageBox-ul de aici
                                    # [System.Windows.Forms.MessageBox]::Show("Biblioteca '$lib' a fost instalată cu succes global.", "Instalare Reușită", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                    $installedApps += "Python Library (Global): $lib"
                                }
                                Remove-Item $tempOutput, $tempError -ErrorAction SilentlyContinue 
                                Write-Host "DEBUG: Cleaned up temporary files for global install."
                                $progressBar.PerformStep() # Pas pentru fiecare bibliotecă
                                $formProgress.Refresh()
                            }
                            # Scoateți MessageBox-ul de aici
                            # [System.Windows.Forms.MessageBox]::Show("Instalare biblioteci Python globale finalizată.", "Succes", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        }
                        else {
                            Write-Host "DEBUG: Pip executable for global install NOT found at: '$globalPipPath'"
                            # Scoateți MessageBox-ul de aici
                            # [System.Windows.Forms.MessageBox]::Show("Pip nu a fost găsit pentru instalare globală. Asigură-te că Python 3 este instalat și căile sunt în PATH.", "Eroare Pip Negăsit", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                            $failedApps += "Python Libraries (Pip negăsit global)"
                            foreach ($lib in $selectedPythonLibraries) { $progressBar.PerformStep(); $formProgress.Refresh() } # Marchează pașii pentru bibliotecile care ar fi trebuit instalate
                        }
                    }
                    catch {
                        Write-Error "ERROR in global pip installation block: $_"
                        # Scoateți MessageBox-ul de aici
                        # [System.Windows.Forms.MessageBox]::Show("A apărut o eroare la rularea pip global: $_", "Eroare Sistem", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        $failedApps += "Python Libraries (Eroare la instalare globală: $($_.Exception.Message))"
                        foreach ($lib in $selectedPythonLibraries) { $progressBar.PerformStep(); $formProgress.Refresh() } # Marchează pașii
                    }

                }
                elseif ($envType -eq "PythonEnv_Virtual") {
                    # Scoateți MessageBox-ul de aici
                    # [System.Windows.Forms.MessageBox]::Show("Instalare biblioteci Python într-un Mediu Virtual a început.", "Instalare Python VEnv", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
               
                    $venvBaseDir = Join-Path $env:USERPROFILE "Documents\PythonVenvs" 
                    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
                    $venvName = "my_app_venv_$timestamp" 
                    $fullVenvPath = Join-Path $venvBaseDir $venvName
                    Write-Host "DEBUG: Virtual environment base directory: '$venvBaseDir'"
                    Write-Host "DEBUG: Virtual environment name: '$venvName'"
                    Write-Host "DEBUG: Full virtual environment path: '$fullVenvPath'"

                    try {
                        if (-not (Test-Path $venvBaseDir)) {
                            New-Item -ItemType Directory -Path $venvBaseDir -Force | Out-Null
                            Write-Host "DEBUG: Created base directory for virtual environments: '$venvBaseDir'"
                        }

                        $labelProgress.Text = "Creare mediu virtual Python..."
                        $formProgress.Refresh()

                        Write-Host "DEBUG: Creating virtual environment at: '$fullVenvPath' using '$pythonExecutable'"
                   
                        $tempVenvOutput = [System.IO.Path]::GetTempFileName()
                        $tempVenvError = [System.IO.Path]::GetTempFileName()

                        $venvProcess = Start-Process -FilePath $pythonExecutable -ArgumentList "-m venv `"$fullVenvPath`"" -NoNewWindow -Wait -PassThru -RedirectStandardOutput $tempVenvOutput -RedirectStandardError $tempVenvError

                        if ($venvProcess.ExitCode -ne 0) {
                            $errorOutput = Get-Content $tempVenvError | Out-String 
                            Write-Error "Failed to create virtual environment. Error: $errorOutput"
                            # Scoateți MessageBox-ul de aici
                            # [System.Windows.Forms.MessageBox]::Show("Eroare la crearea mediului virtual. Vezi consola PowerShell pentru detalii." + "`n" + $errorOutput, "Eroare VEnv", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                            $failedApps += "Python Virtual Environment (creare eșuată)"
                            Remove-Item $tempVenvOutput, $tempVenvError -ErrorAction SilentlyContinue
                            foreach ($lib in $selectedPythonLibraries) { $progressBar.PerformStep(); $formProgress.Refresh() } # Marchează pașii
                            return
                        }
                        else {
                            Write-Host "DEBUG: Virtual environment created successfully at '$fullVenvPath'."
                            $venvOutputContent = Get-Content $tempVenvOutput | Out-String 
                            Write-Host "DEBUG: Virtual environment creation output: $venvOutputContent"
                        }
                        Remove-Item $tempVenvOutput, $tempVenvError -ErrorAction SilentlyContinue
                        Write-Host "DEBUG: Cleaned up temporary files for venv creation."


                        $venvPipPath = Join-Path $fullVenvPath "Scripts\pip.exe"
                        Write-Host "DEBUG: Virtual environment Pip path calculated: '$venvPipPath'"
                        if (-not (Test-Path $venvPipPath)) {
                            Write-Error "ERROR: Pip executable not found in the virtual environment at expected path: '$venvPipPath'"
                            # Scoateți MessageBox-ul de aici
                            # [System.Windows.Forms.MessageBox]::Show("Pip executable not found in the virtual environment. Virtual environment might be corrupted. Check path: `n$venvPipPath", "Eroare Pip VEnv", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                            $failedApps += "Python Virtual Environment (Pip negăsit)"
                            foreach ($lib in $selectedPythonLibraries) { $progressBar.PerformStep(); $formProgress.Refresh() } # Marchează pașii
                            return
                        }
                        else {
                            Write-Host "DEBUG: Pip executable found in virtual environment at: '$venvPipPath'"
                        }

                        foreach ($lib in $selectedPythonLibraries) {
                            $labelProgress.Text = "Instalare bibliotecă Python (VEnv): $lib..."
                            $formProgress.Refresh()

                            Write-Host "DEBUG: Installing '$lib' into virtual environment using '$venvPipPath'..."
                            $tempLibOutput = [System.IO.Path]::GetTempFileName()
                            $tempLibError = [System.IO.Path]::GetTempFileName()

                            $libProcess = Start-Process -FilePath $venvPipPath -ArgumentList "install $lib" -NoNewWindow -Wait -PassThru -RedirectStandardOutput $tempLibOutput -RedirectStandardError $tempLibError

                            if ($libProcess.ExitCode -ne 0) {
                                $errorOutput = Get-Content $tempLibError | Out-String 
                                Write-Error "Failed to install '$lib' in virtual environment. Error: $errorOutput"
                                # Scoateți MessageBox-ul de aici
                                # [System.Windows.Forms.MessageBox]::Show("Eroare la instalarea '$lib' în mediul virtual.`nDetalii: " + $errorOutput, "Eroare Instalare", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                                $failedApps += "Python Library (VEnv): $lib"
                            }
                            else {
                                $outputContent = Get-Content $tempLibOutput | Out-String 
                                Write-Host "Successfully installed '$lib' in virtual environment. Output: $outputContent"
                                # Scoateți MessageBox-ul de aici
                                # [System.Windows.Forms.MessageBox]::Show("Biblioteca '$lib' a fost instalată cu succes în mediul virtual.", "Instalare Reușită", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                $installedApps += "Python Library (VEnv): $lib"
                            }
                            Remove-Item $tempLibOutput, $tempLibError -ErrorAction SilentlyContinue 
                            Write-Host "DEBUG: Cleaned up temporary files for library install."
                            $progressBar.PerformStep() # Pas pentru fiecare bibliotecă
                            $formProgress.Refresh()
                        }
                        # Scoateți MessageBox-ul de aici
                        # [System.Windows.Forms.MessageBox]::Show("Instalare biblioteci Python în mediul virtual '$venvName' finalizată.", "Succes", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        Write-Host "Mediul virtual este localizat la: '$fullVenvPath'"
                    }
                    catch {
                        Write-Error "ERROR in virtual environment management block: $_"
                        # Scoateți MessageBox-ul de aici
                        # [System.Windows.Forms.MessageBox]::Show("A apărut o eroare la gestionarea mediului virtual: $_", "Eroare Sistem", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        $failedApps += "Python Virtual Environment (Eroare internă: $($_.Exception.Message))"
                        foreach ($lib in $selectedPythonLibraries) { $progressBar.PerformStep(); $formProgress.Refresh() } # Marchează pașii
                    }
                }
            }
            else {
                # Scoateți MessageBox-ul de aici
                # [System.Windows.Forms.MessageBox]::Show("Ați selectat Python 3, dar nu ați ales biblioteci de instalat.", "Instalare Python", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                # Nu adăugăm la installed/failed aici, deoarece Python de bază a fost deja marcat.
                Write-Host "INFO: Python 3 selected, but no libraries to install."
                # Dacă Python de bază a fost deja instalat, nu incrementăm progresul pentru biblioteci dacă nu sunt alese.
                # Aceasta este o mică nuanță, dar contează pentru calculul $totalSteps.
            }
        }

        # Aici poți adăuga logica de instalare și pentru celelalte aplicații (winget, etc.)
        # Parcurge $selectedApps (fără Office 2021 și Python 3 care au logică specială)
        foreach ($appText in $selectedApps) {
            $appId = ($appsByCategory.Values | ForEach-Object { $_ } | ForEach-Object { $_ } | Where-Object { $_["Name"] -eq $appText } | ForEach-Object { $_["Id"] })

            Write-Host "DEBUG: App ID for '$appText': $appId"
    
            if ($appId -and $appId -ne "OfficeSetup" -and $appId -notlike "Python.Python.*") {

                $labelProgress.Text = "Verificare $appText..."
                $formProgress.Refresh()

                # --- REVIZUIT: VERIFICARE DACA ESTE DEJA INSTALAT ---
                $isInstalled = $false
                $wingetCheckOutput = ""
                $tempCheckOutput = [System.IO.Path]::GetTempFileName()
                $tempCheckError = [System.IO.Path]::GetTempFileName()

                try {
                    # Run winget list, redirecting output to temporary files
                    $checkProcess = Start-Process -FilePath "winget.exe" `
                        -ArgumentList "list --id $($appId) --exact" `
                        -Wait -NoNewWindow -PassThru `
                        -RedirectStandardOutput $tempCheckOutput `
                        -RedirectStandardError $tempCheckError

                    # Read the content of the temporary output files
                    $wingetCheckOutput = Get-Content $tempCheckOutput | Out-String
                    $wingetCheckError = Get-Content $tempCheckError | Out-String

                    # Check if the process exited successfully and if the output contains the package ID
                    if ($checkProcess.ExitCode -eq 0 -and $wingetCheckOutput -like "*$($appId)*") {
                        Write-Host "DEBUG: '$appText' (ID: $appId) is already installed according to winget list."
                        $isInstalled = $true
                    }
                    elseif ($checkProcess.ExitCode -ne 0 -and $wingetCheckError -like "*No package found matching input criteria*") {
                        # This means winget list explicitly said it's NOT installed
                        Write-Host "DEBUG: '$appText' (ID: $appId) is NOT installed according to winget list."
                        $isInstalled = $false
                    }
                    else {
                        # Other winget list errors or unexpected output
                        Write-Warning "WARNING: Unexpected output from 'winget list $($appId)'. Exit Code: $($checkProcess.ExitCode). Output: '$wingetCheckOutput'. Error: '$wingetCheckError'."
                        # We'll proceed with install attempt if we couldn't confirm it's installed
                        $isInstalled = $false 
                    }
                }
                catch {
                    Write-Warning "WARNING: Exception during 'winget list' check for '$appText'. Error: $($_.Exception.Message). Proceeding with install attempt."
                    # If an exception occurs during the check, assume it's not installed to be safe
                    $isInstalled = $false 
                }
                finally {
                    # Clean up temporary files
                    Remove-Item $tempCheckOutput, $tempCheckError -ErrorAction SilentlyContinue
                }
                # --- SFARSIT REVIZUIRE VERIFICARE ---
        
                if ($isInstalled) {
                    Write-Host "INFO: '$appText' is already installed. Marking as successful and skipping direct installation." -ForegroundColor Green
                    $installedApps += "$appText (Deja instalat)"
                    $progressBar.PerformStep()
                    $formProgress.Refresh()
                    continue # Treci la următoarea aplicație
                }
        
                # --- LOGICA DE INSTALARE (EXISTENTA) ---
                $labelProgress.Text = "Instalare: $appText..."
                $formProgress.Refresh()

                Write-Host "DEBUG: Installing other app via Winget: $appText (ID: $appId)"
                $tempOtherAppOutput = [System.IO.Path]::GetTempFileName()
                $tempOtherAppError = [System.IO.Path]::GetTempFileName()
        
                try {
                    $otherAppProcess = Start-Process -FilePath "winget" -ArgumentList "install --id $($appId) --source winget --accept-package-agreements --accept-source-agreements --silent" -Wait -NoNewWindow -PassThru -RedirectStandardOutput $tempOtherAppOutput -RedirectStandardError $tempOtherAppError
        
                    if ($otherAppProcess.ExitCode -eq 0) {
                        Write-Host "INFO: '$appText' installed successfully!"
                        $installedApps += $appText
                    }
                    else {
                        $otherAppErrorContent = Get-Content $tempOtherAppError | Out-String
                        # Check for the specific "No available upgrade found" message from winget
                        if ($otherAppErrorContent -like "*No available upgrade found*") {
                            Write-Host "INFO: '$appText' is already at the latest version. No upgrade needed." -ForegroundColor Green
                            $installedApps += "$appText (Cea mai recentă versiune)"
                        }
                        else {
                            Write-Error "ERROR: Failed to install '$appText'. Winget Exit Code: $($otherAppProcess.ExitCode). Details: $otherAppErrorContent"
                            $failedApps += "$appText (Winget eșec, cod: $($otherAppProcess.ExitCode))"
                        }
                    }
                }
                catch {
                    Write-Error "ERROR: Exception during '$appText' installation: $_"
                    $failedApps += "$appText (Excepție: $($_.Exception.Message))"
                }
                finally {
                    Remove-Item $tempOtherAppOutput, $tempOtherAppError -ErrorAction SilentlyContinue
                    Write-Host "DEBUG: Cleaned up temporary winget output files for '$appText'."
                }

                $progressBar.PerformStep()
                $formProgress.Refresh()
            }
        }
   
        # Mută aceste linii la finalul întregii logici de instalare
        Start-Sleep -Milliseconds 500
        $formProgress.Close() # Închide bara de progres abia la final

        # === AFISARE REZUMAT FINAL ===
        if ($installedApps.Count -gt 0 -or $failedApps.Count -gt 0) {
            $summary = "Instalare finalizată!"
            if ($installedApps.Count -gt 0) {
                $summary += "`n`nAplicații instalate cu succes:`n" + ($installedApps -join "`n")
            }
            if ($failedApps.Count -gt 0) {
                $summary += "`n`nAplicații care NU au fost instalate:`n" + ($failedApps -join "`n")
            }
            [System.Windows.Forms.MessageBox]::Show($summary, "Rezumat instalare", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Nicio aplicație nu a fost instalată sau nu au fost selectate aplicații.", "Instalare Fără Acțiune", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }

    })





Refresh-AppList # apel inițial

# === TAB 2: Activation & Download ===
$tabActivation = New-Object System.Windows.Forms.TabPage
$tabActivation.Text = "Activation & Tools"
$tabControl.TabPages.Add($tabActivation)

# --- Controls pentru activare ---

$cbActivateWindows = New-Object System.Windows.Forms.CheckBox
$cbActivateWindows.Text = "Activate Windows"
$cbActivateWindows.AutoSize = $true
$cbActivateWindows.Location = New-Object System.Drawing.Point($script:margin, 20)
$tabActivation.Controls.Add($cbActivateWindows)

$cbActivateOffice = New-Object System.Windows.Forms.CheckBox
$cbActivateOffice.Text = "Activate Office"
$cbActivateOffice.AutoSize = $true
$cbActivateOffice.Location = New-Object System.Drawing.Point($script:margin, 50)
$tabActivation.Controls.Add($cbActivateOffice)

$btnActivate = New-Object System.Windows.Forms.Button
$btnActivate.Text = "Run Activation"
$btnActivate.Size = New-Object System.Drawing.Size(150, 30)
$btnActivate.Location = New-Object System.Drawing.Point($script:margin, 90)
$tabActivation.Controls.Add($btnActivate)

$lblActivationStatus = New-Object System.Windows.Forms.Label
$lblActivationStatus.Text = "Status: Waiting..."
$lblActivationStatus.AutoSize = $true
$lblActivationStatus.Location = New-Object System.Drawing.Point($script:margin, 130)
$tabActivation.Controls.Add($lblActivationStatus)

# --- Controls pentru descărcare cursuri ---

$lblUrl = New-Object System.Windows.Forms.Label
$lblUrl.Text = "Custom Courses URL:"
$lblUrl.AutoSize = $true
$lblUrl.Location = New-Object System.Drawing.Point($script:margin, 180)
$tabActivation.Controls.Add($lblUrl)

$txtUrl = New-Object System.Windows.Forms.TextBox
$txtUrl.Size = New-Object System.Drawing.Size(600, 25)
$txtUrl.Location = New-Object System.Drawing.Point($script:margin, 210)
$tabActivation.Controls.Add($txtUrl)

$btnDownload = New-Object System.Windows.Forms.Button
$btnDownload.Text = "Download and Install Courses"
$btnDownload.Size = New-Object System.Drawing.Size(220, 30)
$btnDownload.Location = New-Object System.Drawing.Point($script:margin, 250)
$tabActivation.Controls.Add($btnDownload)

$lblDownloadStatus = New-Object System.Windows.Forms.Label
$lblDownloadStatus.Text = "Status: Waiting..."
$lblDownloadStatus.AutoSize = $true
$lblDownloadStatus.Location = New-Object System.Drawing.Point($script:margin, 290)
$tabActivation.Controls.Add($lblDownloadStatus)

# --- Funcție activare Windows ---
function Activate-Windows {
    try {
        $lblActivationStatus.Text = "Status: Activating Windows..."
        & slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
        & slmgr /skms kms8.msguides.com
        & slmgr /ato
        $lblActivationStatus.Text = "Status: Windows activated (verifică manual)"
    }
    catch {
        $lblActivationStatus.Text = "Status: Activation failed: $_"
    }
}

# --- Funcție activare Office ---
function Activate-Office {
    try {
        $lblActivationStatus.Text = "Status: Activating Office..."
        $officePath = "$env:ProgramFiles\Microsoft Office\Office16\OSPP.VBS"
        if (Test-Path $officePath) {
            cscript.exe $officePath /act
            $lblActivationStatus.Text = "Status: Office activated (verifică manual)"
        }
        else {
            $lblActivationStatus.Text = "Status: Office OSPP.VBS nu găsit"
        }
    }
    catch {
        $lblActivationStatus.Text = "Status: Activation failed: $_"
    }
}

# --- Funcție de descărcare + instalare cursuri ---
function Download-And-Install-Courses {
    param(
        [string]$url
    )
    try {
        $lblDownloadStatus.Text = "Status: Descărcare în curs..."
        $tempPath = Join-Path -Path $env:TEMP -ChildPath "CustomCourses"
        if (-not (Test-Path $tempPath)) { New-Item -ItemType Directory -Path $tempPath | Out-Null }

        $fileName = [System.IO.Path]::GetFileName($url)
        $localFile = Join-Path -Path $tempPath -ChildPath $fileName

        Invoke-WebRequest -Uri $url -OutFile $localFile -UseBasicParsing

        $lblDownloadStatus.Text = "Status: Descărcare completă, instalare..."

        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $extractPath = Join-Path -Path $tempPath -ChildPath "Extracted"
        if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
        [System.IO.Compression.ZipFile]::ExtractToDirectory($localFile, $extractPath)

        $lblDownloadStatus.Text = "Status: Instalare finalizată."
    }
    catch {
        $lblDownloadStatus.Text = "Status: Eroare la descărcare/instalare: $_"
    }
}

# Evenimente butoane activare și descărcare
$btnActivate.Add_Click({
        if ($cbActivateWindows.Checked) {
            Activate-Windows
        }
        if ($cbActivateOffice.Checked) {
            Activate-Office
        }
        if (-not $cbActivateWindows.Checked -and -not $cbActivateOffice.Checked) {
            [System.Windows.Forms.MessageBox]::Show("Selectați cel puțin o opțiune de activare.", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })

$btnDownload.Add_Click({
        $url = $txtUrl.Text.Trim()
        if ([string]::IsNullOrEmpty($url)) {
            [System.Windows.Forms.MessageBox]::Show("Introduceți un URL valid.", "Eroare", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        Download-And-Install-Courses -url $url
    })

# --- RUN FORM ---
[System.Windows.Forms.Application]::Run($form)