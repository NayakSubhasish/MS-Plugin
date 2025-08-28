import * as React from "react";
import { makeStyles, tokens } from "@fluentui/react-components";
import {
  Button,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogContent,
  DialogActions,
  Textarea,
  Label,
  SelectTabData,
  SelectTabEvent,
  Tab,
  TabList,
  TabValue,
} from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  textarea: {
    width: "100%",
    minHeight: "300px",
    fontFamily: "monospace",
    fontSize: "14px",
    padding: "8px",
    borderRadius: "4px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    resize: "vertical",
  },
  tabContent: {
    padding: "16px 0",
  },
});

const defaultPrompts = {
  writeEmail: "Write a professional email with the following details:\nDescription: {description}\nTone: {tone}\nPoint of View: {pointOfView}",
  suggestReply: "Suggest a professional reply to this email considering the specified tone and point of view:\n\nEmail to reply to:\n{emailBody}\n\nResponse Requirements:\nTone: {tone}\nPoint of View: {pointOfView}\n\nPlease craft a reply that matches the specified tone and perspective.",
  summarize: "Summarize this email in 2 sentences:\n{emailBody}",
};

const PromptConfig = ({ onSavePrompts }) => {
  const styles = useStyles();
  const [open, setOpen] = React.useState(false);
  const [selectedTab, setSelectedTab] = React.useState("writeEmail");
  const [prompts, setPrompts] = React.useState(defaultPrompts);

  const handleTabSelect = (event, data) => {
    setSelectedTab(data.value);
  };

  const handlePromptChange = (value) => {
    setPrompts((prev) => ({
      ...prev,
      [selectedTab]: value,
    }));
  };

  const handleSave = () => {
    onSavePrompts(prompts);
    setOpen(false);
  };

  const handleReset = () => {
    setPrompts(defaultPrompts);
  };

  return (
    <Dialog open={open} onOpenChange={(e, data) => setOpen(data.open)}>
      <DialogTrigger>
        <Button 
          appearance="secondary" 
          style={{
            width: '100%',
            minHeight: '48px',
            fontSize: '14px',
            fontWeight: '600',
            borderRadius: '8px',
            border: '1px solid #d1d1d1',
            backgroundColor: '#ffffff',
            color: '#323130',
          }}
        >
          Configure Prompt
        </Button>
      </DialogTrigger>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Configure System Prompt</DialogTitle>
          <DialogContent>
            <div className={styles.root}>
              <TabList selectedValue={selectedTab} onTabSelect={handleTabSelect}>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               <Tab value="writeEmail" title="Write Email">
                       <span style={{ display: 'flex', alignItems: 'center' }}>
                         <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                           <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/>
                           <polyline points="22,6 12,13 2,6"/>
                         </svg>
                       </span>
                     </Tab>
                     <Tab value="suggestReply" title="Suggest Reply">
                       <span style={{ display: 'flex', alignItems: 'center' }}>
                         <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                           <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
                         </svg>
                       </span>
                     </Tab>
                     <Tab value="summarize" title="Summarize">
                       <span style={{ display: 'flex', alignItems: 'center' }}>
                         <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                           <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                           <polyline points="14,2 14,8 20,8"/>
                           <line x1="16" y1="13" x2="8" y2="13"/>
                           <line x1="16" y1="17" x2="8" y2="17"/>
                           <polyline points="10,9 9,9 8,9"/>
                         </svg>
                       </span>
                     </Tab>
              </TabList>
              <div className={styles.tabContent}>
                <Label>System Prompt Template</Label>
                <Textarea
                  className={styles.textarea}
                  value={prompts[selectedTab]}
                  onChange={(e) => handlePromptChange(e.target.value)}
                  placeholder="Enter your custom prompt template..."
                  rows={10}
                />
                <div style={{ fontSize: "12px", color: tokens.colorNeutralForeground3, marginTop: "4px" }}>
                  Available placeholders: {"{emailBody}"} (email content), {"{tone}"} (response tone), {"{pointOfView}"} (perspective), {"{description}"} (for writeEmail)
                </div>
              </div>
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={handleReset}>Reset to Default</Button>
            <Button appearance="primary" onClick={handleSave}>Save Changes</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};

export default PromptConfig; 