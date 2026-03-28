import uno
import unohelper
import json
import urllib.request
import urllib.parse
import os
from com.sun.star.task import XJobExecutor
from com.sun.star.awt import XActionListener

CONFIG_FILE = os.path.expanduser("~/.lmstudio_ext_config.json")

def load_config():
    try:
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    except Exception:
        return {"api_url": "http://127.0.0.1:1234", "model_name": "local-model", "context_words": "300"}

def save_config(api_url, model_name, context_words):
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump({"api_url": api_url, "model_name": model_name, "context_words": context_words}, f)
    except Exception:
        pass

class MyActionListener(unohelper.Base, XActionListener):
    def __init__(self, callback):
        self.callback = callback

    def actionPerformed(self, actionEvent):
        self.callback(actionEvent)

    def disposing(self, source):
        pass

class LMStudioJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx
        self.dialog_control = None
        self.doc = None
        self.target_range = None
        self.current_found = None
        self.last_search_word = ""

    def trigger(self, args):
        try:
            desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
            self.doc = desktop.getCurrentComponent()
            if not self.doc:
                return

            self.create_and_show_dialog()
                
        except Exception as e:
            self.msgbox(f"Trigger Error: {str(e)}")

    def create_and_show_dialog(self):
        smgr = self.ctx.ServiceManager
        
        # 1. Create Model
        dialog_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", self.ctx)
        dialog_model.Title = "Lexico: Find & Replace"
        dialog_model.Width = 300
        dialog_model.Height = 350
        
        # 2. Add Controls
        def add_label(name, label, x, y, w, h):
            model = dialog_model.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
            model.Name = name
            model.Label = label
            model.PositionX = x
            model.PositionY = y
            model.Width = w
            model.Height = h
            dialog_model.insertByName(name, model)
            
        def add_text(name, text, x, y, w, h, multi=False, readonly=False):
            model = dialog_model.createInstance("com.sun.star.awt.UnoControlEditModel")
            model.Name = name
            model.Text = text
            model.PositionX = x
            model.PositionY = y
            model.Width = w
            model.Height = h
            if multi:
                model.MultiLine = True
                model.VScroll = True
            if readonly:
                model.ReadOnly = True
            dialog_model.insertByName(name, model)
            
        def add_button(name, label, x, y, w, h, btn_type=None):
            model = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
            model.Name = name
            model.Label = label
            model.PositionX = x
            model.PositionY = y
            model.Width = w
            model.Height = h
            if btn_type is not None:
                model.PushButtonType = btn_type
            dialog_model.insertByName(name, model)
            
        config = load_config()
        
        # Settings
        add_label("lblApiUrl", "API URL:", 5, 10, 60, 12)
        add_text("txtApiUrl", config.get("api_url", "http://127.0.0.1:1234"), 70, 8, 220, 15)
        
        add_label("lblModelName", "Model:", 5, 30, 60, 12)
        add_text("txtModelName", config.get("model_name", "local-model"), 70, 28, 220, 15)
        
        # Search & Prompt
        add_label("lblWord", "Find Word:", 5, 50, 60, 12)
        add_text("txtWord", "target_word", 70, 48, 120, 15)
        
        add_label("lblContext", "± Words:", 190, 50, 50, 12)
        add_text("txtContext", config.get("context_words", "300"), 240, 48, 50, 15)

        
        add_label("lblPrompt", "Instruction:", 5, 70, 60, 12)
        add_text("txtPrompt", "Rewrite this paragraph to be more concise.", 70, 68, 220, 15)
        
        # Action Buttons
        add_button("btnPrev", "Prev", 70, 90, 50, 20)
        add_button("btnNext", "Next", 125, 90, 50, 20)
        add_button("btnGenerate", "Generate Preview", 180, 90, 110, 20)
        
        # Previews
        add_label("lblOrig", "Original Paragraph:", 5, 120, 100, 12)
        add_text("txtOrig", "", 5, 135, 285, 60, multi=True, readonly=True)
        
        add_label("lblNew", "New Paragraph (Editable):", 5, 200, 130, 12)
        add_text("txtNew", "", 5, 215, 285, 80, multi=True)
        
        # Final Actions
        add_button("btnApprove", "Approve & Replace", 60, 310, 80, 20)
        
        # Cancel button using standard push button type for closing cleanly
        btn_cancel = uno.Enum("com.sun.star.awt.PushButtonType", "CANCEL")
        add_button("btnCancel", "Close", 160, 310, 80, 20, btn_cancel)
        
        # 3. Create Dialog Control
        self.dialog_control = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", self.ctx)
        self.dialog_control.setModel(dialog_model)
        
        # 4. Attach Listeners
        ctrl_prev = self.dialog_control.getControl("btnPrev")
        ctrl_prev.addActionListener(MyActionListener(lambda e: self.on_find(e, True)))
        
        ctrl_next = self.dialog_control.getControl("btnNext")
        ctrl_next.addActionListener(MyActionListener(lambda e: self.on_find(e, False)))
        
        ctrl_gen = self.dialog_control.getControl("btnGenerate")
        ctrl_gen.addActionListener(MyActionListener(self.on_generate))
        
        ctrl_approve = self.dialog_control.getControl("btnApprove")
        ctrl_approve.addActionListener(MyActionListener(self.on_approve))
        
        # 5. Create Peer & Execute
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
        self.dialog_control.createPeer(toolkit, None)
        
        self.dialog_control.execute()
        
        # Save config when closed
        api_url = self.dialog_control.getControl("txtApiUrl").getText()
        model_name = self.dialog_control.getControl("txtModelName").getText()
        context_words = self.dialog_control.getControl("txtContext").getText()
        save_config(api_url, model_name, context_words)
        
        self.dialog_control.dispose()

    def on_find(self, event, backwards):
        try:
            word = self.dialog_control.getControl("txtWord").getText()
            if not word:
                self.msgbox("Please enter a word to find.")
                return
                
            search = self.doc.createSearchDescriptor()
            search.SearchString = word
            search.SearchCaseSensitive = False
            search.SearchBackwards = backwards
            
            # Reset found state if searching for a different word
            if self.last_search_word != word:
                self.current_found = None
            
            if not self.current_found:
                found = self.doc.findFirst(search)
            else:
                start_point = self.current_found.getStart() if backwards else self.current_found.getEnd()
                found = self.doc.findNext(start_point, search)
                
            if not found:
                self.msgbox("No more instances found.")
                return
                
            self.last_search_word = word
            self.current_found = found
            
            # Optionally select the word in the editor to visually guide the user
            controller = self.doc.getCurrentController()
            if controller:
                controller.select(found)
            
            context_words_str = self.dialog_control.getControl("txtContext").getText()
            try:
                context_words = int(context_words_str)
            except:
                context_words = 300
                
            # Get +/- context_words context
            cursor_start = found.getText().createTextCursorByRange(found.getStart())
            for _ in range(context_words):
                if not cursor_start.gotoPreviousWord(False):
                    break
                    
            cursor_end = found.getText().createTextCursorByRange(found.getEnd())
            for _ in range(context_words):
                if not cursor_end.gotoNextWord(False):
                    break
            
            cursor_start.gotoRange(cursor_end, True)
            self.target_range = cursor_start
            
            original_text = cursor_start.getString()
            self.dialog_control.getControl("txtOrig").setText(original_text)
            self.dialog_control.getControl("txtNew").setText("")
            
        except Exception as ex:
            self.msgbox(f"Find Error: {str(ex)}")

    def on_generate(self, event):
        try:
            if not getattr(self, 'target_range', None):
                self.msgbox("No context available. Please click Prev or Next to find the word first.")
                return
                
            prompt_instruction = self.dialog_control.getControl("txtPrompt").getText()
            api_url = self.dialog_control.getControl("txtApiUrl").getText()
            model_name = self.dialog_control.getControl("txtModelName").getText()
            original_text = self.dialog_control.getControl("txtOrig").getText()
            
            if not original_text:
                self.msgbox("Original text is empty.")
                return
                
            self.dialog_control.getControl("txtNew").setText("Generating...")
            
            # API Call
            data = {
                "model": model_name,
                "messages": [
                    {"role": "system", "content": "You are a helpful assistant that edits text based on user instructions. Return ONLY the edited paragraph text without any explanations."},
                    {"role": "user", "content": f"Instruction: {prompt_instruction}\n\nText to edit:\n{original_text}"}
                ],
                "temperature": 0.7
            }
            
            req = urllib.request.Request(f"{api_url}/v1/chat/completions")
            req.add_header('Content-Type', 'application/json')
            
            try:
                response = urllib.request.urlopen(req, json.dumps(data).encode('utf-8'))
                resp_json = json.loads(response.read().decode('utf-8'))
                edited_text = resp_json['choices'][0]['message']['content']
                
                self.dialog_control.getControl("txtNew").setText(edited_text)
                
            except Exception as e:
                self.dialog_control.getControl("txtNew").setText(f"API Error: {str(e)}")
                
        except Exception as ex:
            self.msgbox(f"Generate Error: {str(ex)}")

    def on_approve(self, event):
        try:
            if not self.target_range:
                self.msgbox("No target paragraph selected. Please Find & Generate first.")
                return
                
            new_text = self.dialog_control.getControl("txtNew").getText()
            if not new_text or new_text.startswith("Generating...") or new_text.startswith("API Error:"):
                self.msgbox("Invalid preview text.")
                return
                
            self.target_range.setString(new_text)
            self.msgbox("Text Replaced!")
            
            # Advance the search end immediately after the newly inserted text
            # We must maintain current_found so 'Next'/'Prev' knows where to search from next time.
            # We can use the start of the replaced text as current_found for future reference.
            new_text_cursor = self.target_range.getText().createTextCursorByRange(self.target_range.getStart())
            self.current_found = new_text_cursor
            
            # Reset target range
            self.target_range = None
            self.dialog_control.getControl("txtOrig").setText("")
            self.dialog_control.getControl("txtNew").setText("")
            
        except Exception as e:
            self.msgbox(f"Approve Error: {str(e)}")
    def msgbox(self, message):
        desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        frame = desktop.getCurrentFrame()
        window = frame.getContainerWindow()
        toolkit = window.getToolkit()
        msgbox = toolkit.createMessageBox(window, "infobox", 1, "Lexico", message)
        msgbox.execute()

# Registration
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    LMStudioJob,
    "org.example.lexico.extension",
    ("com.sun.star.task.Job",),)
