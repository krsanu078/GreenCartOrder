set -e

REPO_DIR="${1:-.}"
cd "$REPO_DIR"

# create and switch to branch
git fetch origin
git checkout -B selenium-bdd-setup

# create directories
mkdir -p features/steps pages .github/workflows

# requirements.txt
cat > requirements.txt <<'EOF'
selenium>=4.0.0
behave>=1.2.6
webdriver-manager>=4.0.0
python-dotenv>=0.21.0
