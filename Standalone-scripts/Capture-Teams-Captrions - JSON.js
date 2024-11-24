if (localStorage.getItem("transcripts") !== null) {
    localStorage.removeItem("transcripts");
}

const transcriptArray = JSON.parse(localStorage.getItem("transcripts")) || [];

function checkTranscripts() {
    const closedCaptionsContainer = document.querySelector("[data-tid='closed-captions-renderer']")
    if (!closedCaptionsContainer) {
        alert("Please, click 'More' > 'Language and speech' > 'Turn on life captions'");
        return;
    }
    const transcripts = closedCaptionsContainer.querySelectorAll('.ui-chat__item');

    transcripts.forEach(transcript => {
        const ID = transcript.querySelector('.fui-Flex > .ui-chat__message').id;
        const Name = transcript.querySelector('.ui-chat__message__author').innerText;
        const Text = transcript.querySelector('.fui-StyledText').innerText;
        const Time = new Date().toISOString().replace('T', ' ').slice(0, -1);

        const index = transcriptArray.findIndex(t => t.ID === ID);

        if (index > -1) {
            if (transcriptArray[index].Text !== Text) {
                // Update the transcript if text changed
                transcriptArray[index] = {
                    Name,
                    Text,
                    Time,
                    ID
                };
            }
        } else {
            console.log({
                Name,
                Text,
                Time,
                ID
            });
            // Add new transcript
            transcriptArray.push({
                Name,
                Text,
                Time,
                ID
            });
        }
    });

    localStorage.setItem('transcripts', JSON.stringify(transcriptArray));
}

const observer = new MutationObserver(checkTranscripts);
observer.observe(document.body, {
    childList: true,
    subtree: true
});

setInterval(checkTranscripts, 10000);

// Download JSON
function downloadJSON() {
    let transcripts = JSON.parse(localStorage.getItem('transcripts'));
    // Remove IDs
    transcripts = transcripts.map(({
        ID,
        ...rest
    }) => rest);

    // Stringify with pretty printing
    const prettyTranscripts = JSON.stringify(transcripts, null, 2);

    // Use the page's title as part of the file name, replacing "__Microsoft_Teams" with nothing
    // and removing any non-alphanumeric characters except spaces
    let title = document.title.replace("__Microsoft_Teams", '').replace(/[^a-z0-9 ]/gi, '');
    const fileName = "transcript - " + title.trim() + ".json";

    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(prettyTranscripts);
    const downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute("href", dataStr);
    downloadAnchorNode.setAttribute("download", fileName);
    document.body.appendChild(downloadAnchorNode); // required for firefox
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
}

function formatTranscriptsAsConversation(includeTime = true) {
    const transcripts = JSON.parse(localStorage.getItem('transcripts')) || [];

    if (transcripts.length === 0) {
        console.log("No transcripts available.");
        return;
    }

    let result = '';
    let currentSpeaker = '';
    let startTime = '';
    let endTime = '';
    let conversation = '';

    transcripts.forEach((transcript, index) => {
        const { Name, Text, Time } = transcript;

        // If the speaker changes or it's the last entry
        if (Name !== currentSpeaker || index === transcripts.length - 1) {
            // Append the current speaker's conversation if any
            if (currentSpeaker) {
                if (includeTime) {
                    result += `**${currentSpeaker} [${startTime} - ${endTime}]:**  \n${conversation.trim()}  \n\n`;
                } else {
                    result += `**${currentSpeaker}:**  \n${conversation.trim()}  \n\n`;
                }
            }

            // Reset for the new speaker
            currentSpeaker = Name;
            startTime = Time;
            conversation = '';
        }

        // Append current text to the conversation
        conversation += `${Text}  \n`;

        // Update the endTime to the latest time
        endTime = Time;
    });

    // Append the last speaker's conversation
    if (conversation.trim()) {
        if (includeTime) {
            result += `**${currentSpeaker} [${startTime} - ${endTime}]:**  \n${conversation.trim()}  \n\n`;
        } else {
            result += `**${currentSpeaker}:**  \n${conversation.trim()}  \n\n`;
        }
    }

    console.log(result.trim());
    return result.trim();
}

function downloadConversation(includeTime = true) {
    // Get the formatted conversation
    const conversation = formatTranscriptsAsConversation(includeTime);

    if (!conversation) {
        console.log("No conversation to download.");
        return;
    }

    // Create a Blob with the conversation text
    const blob = new Blob([conversation], { type: 'text/plain' });

    // Generate a filename with the page title and timestamp
    const title = document.title.replace(/[^a-zA-Z0-9 ]/g, '').trim(); // Sanitize title
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-'); // Replace : and . with -
    const fileName = `${title || 'conversation'}-${timestamp}.txt`;

    // Create an anchor element to trigger the download
    const downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.href = URL.createObjectURL(blob);
    downloadAnchorNode.download = fileName;

    // Append the anchor to the DOM, click it, and remove it
    document.body.appendChild(downloadAnchorNode);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
}

// lets Download with time
//downloadConversation(true);

// lets Download without time
//downloadConversation(false);

// Let's download the JSON 
// downloadJSON();
