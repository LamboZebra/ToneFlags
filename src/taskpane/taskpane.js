/* global Office */

let selectedTones = [];

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("ToneFlags taskpane loaded and ready.");
    }
});

function toggleTone(buttonElement, toneName, colorCode) {
    const index = selectedTones.findIndex(t => t.name === toneName);

    if (index > -1) {
        selectedTones.splice(index, 1);
        buttonElement.classList.remove("selected");
    } else {
        selectedTones.push({ name: toneName, color: colorCode });
        buttonElement.classList.add("selected");
    }
}

function insertToneHeader() {
    if (selectedTones.length === 0) {
        showStatus("Select at least one tone flag first.");
        return;
    }

    let badgesHtml = `<div id="toneflag-header" style="font-family: 'Segoe UI', Arial, sans-serif; margin-bottom: 16px; padding: 12px; background-color: #f8fafc; border-left: 4px solid #3b82f6; border-radius: 4px;">`;
    badgesHtml += `<div style="font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase; margin-bottom: 8px;">Tone Flags & Intent:</div>`;
    
    selectedTones.forEach(tone => {
        let description = "The sender would like to convey this specific tone and intent.";
        
        switch (tone.name) {
            case "Rushed":
                description = "The sender needs to be brief due to external time pressures. Please do not read this as blunt or abrupt.";
                break;
            case "Direct":
                description = "Communicating straight to the point for clarity and efficiency. There is no hidden frustration or negative intent.";
                break;
            case "Banter":
                description = "Using lighthearted humor or playful teasing in a friendly spirit. Intended to connect, not criticise.";
                break;
            case "Playful":
                description = "Written in a light, creative, or fun-spirited frame of mind to keep things relaxed.";
                break;
            case "Sarcasm":
                description = "Using ironic or deadpan humor in jest without any insult intent. Not to be taken literally.";
                break;
            case "Serious":
                description = "Requires focused attention on an important matter.";
                break;
            case "Warm":
                description = "Written with friendly, supportive, and positive regard.";
                break;
            case "Reassurance":
                description = "The sender wishes to convey reassurance with this message, offering comfort and/or validation.";
                break;
            case "Encouragement":
                description = "Intending to show belief in your abilities and that you're on the right track.";
                break;
            case "Venting":
                description = "Blowing off some steam or sharing frustration about a situation. This is not directed at you personally and should not be considered an attack.";
                break;
            case "Concerned":
                description = "The sender wishes to convey that they have a degree of worry or anxiety about the topic being discussed and would appreciate reassurance or further information.";
                break;
            case "Idea Only":
                description = "Floating a tentative thought or brainstorm concept. There is zero pressure or final commitment attached.";
                break;
        }

        badgesHtml += `<div style="margin-bottom: 6px;">`;
        badgesHtml += `<span style="display: inline-block; background-color: ${tone.color}; color: #ffffff; font-size: 11px; font-weight: 600; padding: 2px 8px; border-radius: 12px; margin-right: 6px;">${tone.name}</span>`;
        badgesHtml += `<span style="font-size: 12px; color: #334155; line-height: 1.4;">${description}</span>`;
        badgesHtml += `</div>`;
    });
    
    badgesHtml += `</div><br/><br/>`;

    Office.context.mailbox.item.body.prependAsync(
        badgesHtml,
        { coercionType: Office.CoercionType.Html },
        (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                showStatus("Detailed ToneFlags banner added!");
            } else {
                showStatus("Failed to insert banner.");
            }
        }
    );
}

function showStatus(message) {
    const statusDiv = document.getElementById("status");
    if (statusDiv) {
        statusDiv.innerText = message;
        setTimeout(() => { 
            if (statusDiv) statusDiv.innerText = ""; 
        }, 3000);
    } else {
        console.log("Status message:", message);
    }
}