Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  const itemBody = document.getElementById("item-body");
  itemBody.value = "Extracting domain names from selection...";

  const currentItem = Office.context.mailbox.item;

  Office.context.mailbox.getSelectedItemsAsync((result) => {
    let extractedDomains = [];

    // Helper function to extract and clean the domain from any string
    const extractDomain = (str) => {
      if (str && str.includes("@")) {
        // Splits at '@', removes any trailing '>' brackets, and lowercases it
        return str.split("@")[1].replace(">", "").toLowerCase().trim();
      }
      return null;
    };

    // 1. Check active open item (if any)
    if (currentItem) {
      if (currentItem.sender && currentItem.sender.emailAddress) extractedDomains.push(extractDomain(currentItem.sender.emailAddress));
      if (currentItem.from && currentItem.from.emailAddress) extractedDomains.push(extractDomain(currentItem.from.emailAddress));
    }

    // 2. Check the Multi-Select List
    if (result.status === Office.AsyncResultStatus.Succeeded && result.value.length > 0) {
      result.value.forEach(item => {
        let targetString = null;
        
        // Look for the easiest data first
        if (item.sender && item.sender.emailAddress) {
          targetString = item.sender.emailAddress;
        } else if (item.from && item.from.emailAddress) {
          targetString = item.from.emailAddress;
        } else if (item.internetMessageId) {
          // If sender is hidden, grab the Message-ID string instead
          targetString = item.internetMessageId;
        }

        if (targetString) {
          extractedDomains.push(extractDomain(targetString));
        }
      });
    }

    // 3. Vaporize Duplicates and Empty Values
    const uniqueDomains = [...new Set(extractedDomains)]
      .filter(d => d !== null && d !== "")
      .sort();

    // 4. Output Results or Debug Info
    if (uniqueDomains.length > 0) {
      itemBody.value = uniqueDomains.join("\r\n");
    } else {
      // also output some debug diagnostics
      const debugCount = result.value ? result.value.length : 0;
      const firstItem = result.value && result.value[0] ? result.value[0] : currentItem;
      
      itemBody.value = "Failed to find domain names from the selection. See below for details:\n\n" +
                       "Selected sount: " + debugCount + "\n" +
                       "JSON of first item: " + JSON.stringify(firstItem, null, 2);
    }
  });
}