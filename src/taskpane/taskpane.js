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
    let processedIds = new Set();
    
    // Determine if we are in "Bulk Mode"
    // New Outlook often returns the single active item in this array too.
    const isMultiSelect = result.status === Office.AsyncResultStatus.Succeeded && result.value.length > 1;

    /**
     * Helper: extractDomain
     * @param {string} str - The raw string to parse.
     * @param {boolean} allowTechnicalFallback - Whether to accept Message-ID domains.
     */
    const extractDomain = (str, allowTechnicalFallback) => {
      if (!str || !str.includes("@")) return null;

      let domain = str.split("@")[1].replace(">", "").toLowerCase().trim();

      /**
       * HISTORICAL NOTE: THE INFRASTRUCTURE HEURISTIC
       * If we are forced to use a Message-ID (only allowed in Multi-Select), 
       * we apply strict filters to avoid internal server hostnames.
       */
      if (allowTechnicalFallback) {
        const infraPatterns = [/\d+/, /mta/i, /smtp/i, /prod/i, /cluster/i, /local/i, /internal/i];
        const isTechnical = infraPatterns.some(p => p.test(domain)) || (domain.match(/\./g) || []).length > 2;
        
        if (isTechnical) return null;
      }

      return domain;
    };

    /**
     * STEP 1: THE ACTIVE ITEM
     * If only one email is selected, we operate in "Strict Mode". 
     * We ignore the internetMessageId entirely to ensure 100% brand accuracy.
     */
    const activeItem = Office.context.mailbox.item;
    if (activeItem) {
      const id = activeItem.itemId || activeItem.internetMessageId;
      processedIds.add(id);

      let source = (activeItem.sender && activeItem.sender.emailAddress) || 
                   (activeItem.from && activeItem.from.emailAddress);
      
      // Pass 'false' for allowTechnicalFallback if it's the only item
      const domain = extractDomain(source, false); 
      if (domain) extractedDomains.push(domain);
    }

    /**
     * STEP 2: THE SELECTION LIST
     * If multiple emails are selected, we enable the "Safety Net".
     * Because New Outlook's selection API is "shallow," the Message-ID fallback 
     * is necessary to ensure we don't return an empty list for bulk tasks.
     */
    if (result.status === Office.AsyncResultStatus.Succeeded && result.value.length > 0) {
      result.value.forEach(item => {
        const id = item.itemId || item.internetMessageId;
        if (processedIds.has(id)) return;
        processedIds.add(id);

        let bestSource = null;
        let useMessageIdAsFallback = false;

        if (item.sender && item.sender.emailAddress) {
          bestSource = item.sender.emailAddress;
        } else if (item.from && item.from.emailAddress) {
          bestSource = item.from.emailAddress;
        } else if (isMultiSelect && item.internetMessageId) {
          // Only allow Message-ID if we are in Multi-Select mode
          bestSource = item.internetMessageId;
          useMessageIdAsFallback = true;
        }

        const domain = extractDomain(bestSource, useMessageIdAsFallback);
        if (domain) extractedDomains.push(domain);
      });
    }

    const uniqueDomains = [...new Set(extractedDomains)].sort();
    itemBody.value = uniqueDomains.length > 0 ? uniqueDomains.join("\r\n") : "No valid domains found.";
  });
}
