# VBA-Sync CI/CD Build System

## Overview
This system automatically compiles your VBA source files (.bas/.cls) into a distributable .xlam add-in file whenever you push a tag to GitHub.

## How It Works

### 1. **vbaProject-Compiler** (MS-OVBA)
- **Purpose**: Compiles individual VBA source files into vbaProject.bin
- **Input**: .bas, .cls files from your VBA Sync folder structure
- **Output**: vbaProject.bin (compiled VBA binary)
- **Repo**: https://github.com/Beakerboy/MS-OVBA

### 2. **Excel-Addin-Generator**
- **Purpose**: Wraps vbaProject.bin into a valid .xlam file
- **Input**: vbaProject.bin
- **Output**: VBA Sync.xlam (distributable add-in)
- **Repo**: https://github.com/James-McNeil/Excel-Addin-Generator (your fork)

### 3. **build_xlam.py** (Custom Build Script)
- **Purpose**: Orchestrates the two tools above
- **Process**:
  1. Scans your VBA Sync folder for all .bas/.cls files
  2. Uses vbaProject-Compiler to build vbaProject.bin
  3. Uses Excel-Addin-Generator to wrap it into .xlam
  4. Cleans up temporary files

## Workflow

\\\
VBA Source Files  vbaProject-Compiler  vbaProject.bin  Excel-Addin-Generator  VBA Sync.xlam
\\\

## GitHub Actions Workflow

When you push a tag (e.g., v1.0.0):
1. **Checkout**: Gets your code from GitHub
2. **Setup Python**: Installs Python 3.x
3. **Install Tools**: Installs both compiler tools
4. **Build**: Runs build_xlam.py to compile your source
5. **Release**: Creates a GitHub Release with the .xlam file attached

## Usage

### Local Development
1. Make changes to VBA code in Excel
2. Run VBA Sync Export to save to source files
3. Commit and push changes to GitHub
4. Create and push a tag to trigger build

\\\ash
git tag v1.0.0
git push origin v1.0.0
\\\

### Automatic Build
- GitHub Actions automatically builds and releases the .xlam
- Users can download the latest version from GitHub Releases

## Version Control with Semantic Versioning

The tag-version.yml workflow automatically:
- Reads your latest tag
- Increments the version based on commit messages:
  - #major  v1.0.0 to v2.0.0
  - #minor  v1.0.0 to v1.1.0  
  - (default)  v1.0.0 to v1.0.1 (patch)
- Creates a new tag
- Triggers the deploy workflow

## Benefits

 **No Excel Required**: Builds run on Linux runners (no Windows/Excel needed)
 **True CI/CD**: Fully automated from commit to release
 **Version Control**: All VBA code in Git as plain text
 **Collaboration**: Team members can review code changes
 **AI Assistance**: AI tools can read and modify VBA source files
 **Reproducible Builds**: Same source always produces same output

## Files in This System

- **build_xlam.py**: Main build script
- **.github/workflows/deploy-xlam.yml**: Release workflow (triggered by tags)
- **.github/workflows/tag-version.yml**: Auto-versioning workflow (triggered by push to main)
- **vba-sync/**: Your VBA source files organized by type

## Next Steps

1. Test the build locally (install dependencies and run build_xlam.py)
2. Commit build_xlam.py and workflow files
3. Push to main to test auto-versioning
4. Verify the .xlam file builds correctly
5. Download and test the .xlam from GitHub Releases

## Troubleshooting

If the build fails, check:
- All .bas/.cls files are properly formatted
- The vba-sync folder structure is correct (Modules/, Objects/, etc.)
- Python dependencies are installed correctly
- GitHub Actions logs for detailed error messages
