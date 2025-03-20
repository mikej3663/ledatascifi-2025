## VS Code Installation

To download VS Code, visit the official [website](https://code.visualstudio.com/) and follow the installation instructions for your operating system (Windows, macOS, or Linux).

## Extensions

You will want to install some extensions. Open VS Code and type `CTRL+SHIFT+X` to open the extensions pane, then search for and install these:

- **Python**: Essential for Python development.
- **Pylance**: Fast, feature-rich language support for Python.
- **Jupyter**: Provides Jupyter Notebook support.
- **Jupyter Keymap**: Adds Jupyter Notebook keybindings.
- **Code Spell Checker**: Checks spelling in multiple languages, including LaTeX.
- **Data Wrangler**: Generates Pandas code for data manipulation.
- **Rainbow CSV**: Helps view and edit CSV files.
- **indent-rainbow**: Visualizes code indentations with colors.
- **Markdown All in One**: Offers various Markdown editing features.
- **Python Indent**: Improves Python indentation.
- **GitHub Copilot**: AI-powered code completions.
- **GitHub Copilot Chat**: Chat interface for GitHub Copilot.
- **VS Code Speech**: Voice control for coding.
- **Path Intellisense**: Autocompletes filenames, and does relative paths. 

## Setting Up Your Workspace

To set your "workspace", the directory you're working within, by  
1. Click on the “File” menu in the top-left corner.
2. Select “Open Folder…” from the dropdown menu.
3. Choose the folder you want to open as your workspace and click “Select Folder.”

After that, you can navigate within that workspace using the explorer tab. A few tricks and comments:
- You can right click on a file and hit "Reveal in File Explorer" to see the file
- You can right click a tab for any open file and "Reveal in Explorer View" will show you it within VS Code
- If you change your workspace location, the files will close and you'll go to another location. If you come back to the same workspace, the files you had open last time will be open.
- **CTRL+I** opens inline help from copilot if you have it installed.
- `CTRL+SPACE` and `CTRL+SHIFT+SPACE` shows parameter hints within functions. (Like `SHIFT+TAB` does in Jupyter Lab.)
- Multi-Cursor Editing: Like some other IDEs, VS Code also has a multi-cursor feature. You can edit multiple lines at once by holding down the Alt (Windows/Linux) or Option (Mac) key and clicking on the lines you want to edit. Here is an example of how powerful this can be: https://x.com/lukestein/status/1444694514338238471

_Note: For Mac users, any time you see CTRL as a shortcut key on this page, replace with CMD._

## Data Analysis

To create a Jupyter Notebook:

1. Open the command palette by pressing Ctrl+Shift+P (Windows/Linux) or Cmd+Shift+P (macOS). **Using the command palette is essential to using VS Code well.**
1. Type “Jupyter: Create New Blank Notebook” and select it. (You can also do File > New File, and then name it with the `ipynb` extension.)
1. Choose the Python interpreter you want to use for the notebook. Look for the "Select Kernel" button in the upper right. VS Code will see any conda environments you have.

To view data/variables:

1. Install the Data Wrangler extension. 
2. Run:

    ```python
    import seaborn as sns
    tips = sns.load_dataset("tips")
    ```

3. Click on “View data” or "Jupyter variables" and choose `tips`. The Data Wrangler will show you the type of each variable (a # for numeric and a 🄰 for objects) and a quick histogram of the variable.
 
## Gen AI and Copilot

GitHub Copilot is excellent for code completions and fixes. Additionally, the chat box almost eliminates the need to open a new window for GenAI use. Note that using GitHub Copilot requires a GitHub Pro account.

**CTRL+I** opens inline help. 

- One great feature of Copilot is that you can open a Jupyter notebook and ask for inline help by pressing Cmd+I. For example, you can type a query like “How to run a Mann-Whitney U test?” and you will see a step-by-step guide appear in your cell.
- Another advantage is that you are not limited to GPT models. In Copilot, you can choose which model to use.
- Copilot chat is like what you're used to using
- Copilot Edits will use the files in your workspace and propose wholesale edits given your direction
- These still don't replace good data science habits, conceptual understanding, and domain expertise. My recent uses of it still require HEAVY edits and interactive feedback in more than half my uses.

## References

Chunks of the above are liberally pulled from Luek Stein, Paul Goldsmith-Pinkham, [Esma Özer](https://esmaozer.github.io/resources/all_in_one/all-in-one/).