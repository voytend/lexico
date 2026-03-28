# Lexico (LibreOffice AI Assistant)

Lexico is a powerful, locally-run AI assistant natively integrated into LibreOffice Writer. Built with Python and the UNO API, it allows you to dynamically edit document context using any OpenAI-compatible Language Model API, such as [LM Studio](https://lmstudio.ai/).

## Features

- **Programmatic Native UI**: A fully dynamic LibreOffice dialog box that avoids common XML UI parser issues.
- **Adjustable Context Windows**: Find any target word and extract an exact number of surrounding words (e.g., ± 300 words) to feed complete context to your AI model.
- **Stateful Search Loop**: Seamlessly cycle through all occurrences of a word in your document using dedicated **Next** and **Prev** buttons.
- **Draft & Preview**: The AI’s output is loaded into an editable preview window so you can make manual tweaks before committing changes.
- **One-Click Replacement**: Approve the generated draft and the original text in the document is instantly replaced.
- **Persistent Settings**: Your API URL, Model Name, and Context Window size are automatically saved across LibreOffice sessions in `~/.lmstudio_ext_config.json`.

## How It Works

1. **Invoke the Dialog**: Go to the **Lexico** menu item in LibreOffice Writer and select **Edit with AI...**.
2. **Configure Connection**: Point the `API URL` to your local LLM server (defaults to `http://127.0.0.1:1234`).
3. **Target Text**: 
   - Specify a `Find Word`.
   - Set the `± Words` boundaries to extract the precise amount of surrounding context.
4. **Locate Instances**: 
   - Click **Next** or **Prev** to locate instances of the word. 
   - The document cursor will jump to the word, and the surrounding text is pulled into the **Original Paragraph** pane.
5. **Generate Rewrite**:
   - Provide an `Instruction` (e.g., *"Rewrite this to be more professional"*).
   - Click **Generate Preview**. The extension sends the instruction and the original paragraph to the AI.
6. **Approve**: View the response in the **New Paragraph** pane. Click **Approve & Replace** to swap the old text within LibreOffice with the new draft.

## Technical Details

- `python/main.py`: Contains the core `XJobExecutor` class that intercepts the menu click, builds the `UnoControlDialog` dialog models programmatically, handles XActionListeners for asynchronous logic, and performs PyUNO document string manipulation.
- `metadata` (`description.xml`, `manifest.xml`, `Addons.xcu`): Standard LibreOffice Extension boilerplate mapping the menu item to the Python UNO component (`org.example.lexico.extension`).
- `package.py`: Utility script to zip the relevant component files into a ready-to-install `.oxt` package.

## Installation

1. Run `python package.py` from the root directory if you need to build the `.oxt` file manually.
2. Open LibreOffice Writer.
3. Navigate to **Tools > Extension Manager**.
4. Click **Add** and select the built `Lexico.oxt` file.
5. Restart LibreOffice.
