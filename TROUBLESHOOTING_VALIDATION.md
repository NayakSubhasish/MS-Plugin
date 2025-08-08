# 🔧 Email Validation Troubleshooting Guide

## 🚨 Issue: Validation Dialog Only Appears Once

**Problem:** The validation dialog appears only once and doesn't reappear for subsequent Send button clicks.

## 🔍 Debug Steps

### **Step 1: Check Console Logs**
Open browser developer tools (F12) and look for these log messages:

```
🚀 Email validation triggered - OnMessageSend event fired
📅 Timestamp: 2024-01-XX...
📧 Email body length: XXX
🔍 Missing keywords found: X
❌ Missing categories: [list]
🔍 Opening validation dialog: https://...
✅ Dialog opened successfully
📤 Sending data to dialog: [missing keywords]
📥 Dialog response received: {"action":"cancel"}
❌ User cancelled
🔒 Closing dialog
```

### **Step 2: Test Manual Validation**
In the browser console, run:
```javascript
testValidation()
```

**Expected:** Should see validation logs and dialog appear.

### **Step 3: Check Dialog URL**
Verify the dialog URL is accessible:
- Open: `https://nayaksubhasish.github.io/MAIL-MAGIC/validation-dialog.html`
- Should show the validation dialog interface

### **Step 4: Test Different Email Scenarios**

#### **Test A: Complete Email (Should Pass)**
```
Dear John,

I hope this email finds you well. I wanted to discuss our upcoming meeting agenda.

Topics to discuss:
- Project timeline and deadlines
- Action items for next steps
- Participants and attendees

Please let me know if you have any questions.

Best regards,
Sarah
```

**Expected:** No dialog, email sends immediately.

#### **Test B: Incomplete Email (Should Show Dialog)**
```
Hi,

Meeting tomorrow at 2pm.

Thanks
```

**Expected:** Dialog appears with missing elements.

#### **Test C: After Cancelling (Should Show Dialog Again)**
1. Write incomplete email
2. Click Send → Dialog appears
3. Click Cancel → Dialog closes
4. Click Send again → Dialog should reappear

## 🛠️ Potential Solutions

### **Solution 1: Clear Browser Cache**
1. Open browser developer tools (F12)
2. Right-click refresh button → "Empty Cache and Hard Reload"
3. Test validation again

### **Solution 2: Check Add-in Registration**
1. In Outlook, go to **Get Add-ins** → **My Add-ins**
2. Remove the Mail Magic add-in
3. Re-add the add-in using `manifest.xml`
4. Test validation again

### **Solution 3: Verify Manifest Configuration**
Check that `manifest.xml` contains:
```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateEmail"/>
</ExtensionPoint>
```

### **Solution 4: Test in Different Scenarios**
- **New Email:** Compose new email → Click Send
- **Draft Email:** Open draft → Click Send  
- **Reply:** Reply to email → Click Send
- **Forward:** Forward email → Click Send

## 🔧 Technical Debugging

### **Check Event Registration**
In browser console, run:
```javascript
console.log('validateEmail function:', typeof validateEmail);
console.log('Office context:', Office.context);
console.log('Mailbox item:', Office.context.mailbox.item);
```

### **Check Dialog Functionality**
In browser console, run:
```javascript
// Test dialog URL
fetch('https://nayaksubhasish.github.io/MAIL-MAGIC/validation-dialog.html')
  .then(response => console.log('Dialog accessible:', response.ok))
  .catch(error => console.error('Dialog not accessible:', error));
```

### **Monitor Network Requests**
1. Open browser developer tools
2. Go to Network tab
3. Click Send button
4. Look for requests to validation-dialog.html

## 🎯 Expected Behavior

### **Correct Flow:**
1. User clicks Send → Validation triggers
2. System checks email content
3. If missing keywords → Dialog appears
4. User chooses: Send Anyway / Add Keywords / Cancel
5. If Cancel → Dialog closes, email doesn't send
6. User clicks Send again → Dialog should reappear

### **Debug Information Added:**
- ✅ **Timestamp logging** for each validation attempt
- ✅ **Dialog URL logging** to verify accessibility
- ✅ **Event completion logging** to track flow
- ✅ **Timeout fallback** to prevent hanging
- ✅ **Manual test function** for debugging

## 🚀 Quick Fixes

### **If Dialog Not Appearing:**
1. Check console for error messages
2. Verify dialog URL is accessible
3. Clear browser cache
4. Reinstall add-in

### **If Validation Not Triggering:**
1. Check manifest.xml has ItemSend event
2. Verify add-in is properly loaded
3. Test with manual validation function

### **If Dialog Appears But Doesn't Work:**
1. Check dialog response handling
2. Verify Office.context.ui.displayDialogAsync
3. Test dialog URL directly in browser

## 📞 Support Information

### **Console Logs to Share:**
```
🚀 Email validation triggered - OnMessageSend event fired
📅 Timestamp: [timestamp]
📧 Email body length: [number]
🔍 Missing keywords found: [number]
❌ Missing categories: [list]
🔍 Opening validation dialog: [URL]
✅ Dialog opened successfully
📤 Sending data to dialog: [data]
📥 Dialog response received: [response]
```

### **Files to Check:**
- `manifest.xml` - Event registration
- `src/commands/commands.js` - Validation logic
- `validation-dialog.html` - Dialog interface
- Browser console - Error messages

---

**🎯 With the enhanced debugging, you should now see detailed logs that will help identify exactly where the issue is occurring!** 