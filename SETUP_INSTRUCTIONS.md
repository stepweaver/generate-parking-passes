### Step 1: Copy Files to Computer
1. Create a folder where you want to install the script
2. Copy all these files/folders from your computer to theirs:
   - `src` folder
   - `credentials` folder
   - `templates` folder
   - `assets` folder
   - `.env` file
   - `requirements.txt`

### Step 2: Install Python Dependencies
1. Open a command prompt/terminal computer
2. Navigate to the script folder
3. Run: `pip install -r requirements.txt`

### Step 3: Configure the .env File
1. Open the `.env` file in a text editor
2. Update the `GMAIL_DELEGATE_EMAIL` to the new delegate email address

### Step 4: Initial Authentication
1. Delete the existing `token.pickle` file from the `credentials` folder (if it exists)
2. Run the script: `python src/generate_guest_passes.py`
3. A browser window will open asking them to sign in with their Google account
4. They should sign in with their ND account that has delegate access
5. After authentication, a new `token.pickle` file will be created in the credentials folder

### Step 5: Create Desktop Shortcut
1. Right-click on the desktop
2. Select New > Shortcut
3. For the location, enter:
   ```
   python "FULL_PATH_TO_SCRIPT\src\generate_guest_passes.py"
   ```
   (Replace FULL_PATH_TO_SCRIPT with the actual path where you installed the script)
4. Name the shortcut something like "Parking Pass Generator"

After these steps, your colleague should be able to:
1. Double-click the desktop shortcut to run the script
2. Send emails from their delegate account without needing to re-authenticate
3. The script will use their delegate email permissions

Note: The authentication (token.pickle) will eventually expire (usually after several weeks or months). When this happens, they'll need to re-authenticate by following Step 4 again.