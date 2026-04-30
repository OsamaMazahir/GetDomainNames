Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  const itemBody = document.getElementById("item-body");
  itemBody.value = "Extracting domains...";

  Office.context.mailbox.getSelectedItemsAsync((result) => {
    let extractedDomains = [];

    // Helper to clean the domain
    const extractDomain = (str) => {
      if (str && str.includes("@")) {
        let domain = str.split("@")[1].replace(">", "").toLowerCase().trim();
        // Ignore internal technical hostnames
        if (domain.endsWith(".local") || domain.endsWith(".internal") || domain.endsWith(".lan")) {
          return null;
        }
        return domain;
      }
      return null;
    };

    // Combine current item (if open) and selected items into one list
    let allItems = [];
    if (Office.context.mailbox.item) allItems.push(Office.context.mailbox.item);
    if (result.status === Office.AsyncResultStatus.Succeeded && result.value.length > 0) {
      allItems = allItems.concat(result.value);
    }

    allItems.forEach(item => {
      let bestSource = null;

      // PRIORITY 1: The actual Sender
      if (item.sender && item.sender.emailAddress) {
        bestSource = item.sender.emailAddress;
      } 
      // PRIORITY 2: The "From" field
      else if (item.from && item.from.emailAddress) {
        bestSource = item.from.emailAddress;
      } 
      // PRIORITY 3: Fallback to Message-ID only if the above are missing
      else if (item.internetMessageId) {
        bestSource = item.internetMessageId;
      }

      if (bestSource) {
        const domain = extractDomain(bestSource);
        if (domain) extractedDomains.push(domain);
      }
    });

    // Vaporize duplicates
    const uniqueDomains = [...new Set(extractedDomains)].sort();

    if (uniqueDomains.length > 0) {
      itemBody.value = uniqueDomains.join("\r\n");
    } else {
      // Keep your debug logic just in case
      const firstItem = allItems[0] || {};
      itemBody.value = "Failed to find domain names. See below for details:\n\n" +
                       "JSON Dump: " + JSON.stringify(firstItem, null, 2);
    }
  });
}