[CmdletBinding()]
param (
    [Parameter(HelpMessage = 'The path to the plugin folder (the folder Obsidian loads the plugin from).')]
    [String]
    $PluginFolder = $env:OBSIDIAN_PLUGIN_FOLDER,

    # Legacy parameter: if caller passes -ObsidianPluginRoot + optional -PluginFolderName,
    # derive $PluginFolder from them for backwards compatibility.
    [Parameter(HelpMessage = 'Legacy: path to the .obsidian/plugins directory.')]
    [String]
    $ObsidianPluginRoot = $env:OBSIDIAN_PLUGIN_ROOT,
    [Parameter(HelpMessage = 'Legacy: subfolder name inside ObsidianPluginRoot.')]
    [String]
    $PluginFolderName = 'obsidian-tasks-plugin'
)

# Resolve target folder
if (-not $PluginFolder) {
    if ($ObsidianPluginRoot) {
        $PluginFolder = Join-Path $ObsidianPluginRoot $PluginFolderName
    } else {
        Write-Error "Provide -PluginFolder <path> or set OBSIDIAN_PLUGIN_FOLDER env var."
        return
    }
}

$repoRoot = (Resolve-Path -Path $(git rev-parse --show-toplevel)).Path

# Create the plugin folder if it does not exist yet
if (-not (Test-Path $PluginFolder)) {
    Write-Host "Plugin folder not found — creating: $PluginFolder"
    New-Item -ItemType Directory -Path $PluginFolder -Force | Out-Null
} else {
    Write-Host "Plugin folder found: $PluginFolder"
}

Push-Location $repoRoot
Write-Host "Repo root: $repoRoot"

yarn run build:dev

if ($?) {
    Write-Output 'Build successful'

    $filesToLink = @('main.js', 'styles.css', 'manifest.json')

    foreach ($file in $filesToLink) {
        $target = Join-Path $PluginFolder $file
        $source = Join-Path $repoRoot $file

        # Remove existing file or link
        if (Test-Path $target) {
            Remove-Item $target -Force
        }

        # Try hard link first (no admin required); fall back to copy
        try {
            New-Item -ItemType HardLink -Path $target -Target $source -ErrorAction Stop | Out-Null
            Write-Output "Hard-linked $file"
        } catch {
            Copy-Item $source $target -Force
            Write-Output "Copied $file (hard link failed: $_)"
        }
    }

    $hotreload = Join-Path $PluginFolder '.hotreload'
    if (-not (Test-Path $hotreload)) {
        Write-Output 'Creating .hotreload file'
        '' | Set-Content $hotreload
    }

    yarn run dev

} else {
    Write-Error 'Build failed'
}

Pop-Location
