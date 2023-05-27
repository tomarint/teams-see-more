const contextMenuId: string = "teams-see-more-context-menu-id";
const messageName: string = "teams-see-more-message";
const successMessage: string = "success";

function logError(location: string): void {
  if (chrome.runtime.lastError) {
    console.log(`Error in ${location}: ${chrome.runtime.lastError.message}`);
  }
}

function createContextMenu(): void {
  chrome.contextMenus.remove(contextMenuId, () => {
    if (chrome.runtime.lastError) {
      // This is expected if the context menu has not yet been created.
      // logError("contextMenus.remove");
    }
    chrome.contextMenus.create(
      {
        id: contextMenuId,
        title: chrome.i18n.getMessage("extName"),
        contexts: ["all"],
        documentUrlPatterns: ["https://teams.microsoft.com/*"],
      },
      () => {
        logError("contextMenus.create");
      }
    );
  });
}

chrome.runtime.onInstalled.addListener(createContextMenu);
chrome.runtime.onStartup.addListener(createContextMenu);

chrome.runtime.onUpdateAvailable.addListener((details: chrome.runtime.UpdateAvailableDetails) => {
  console.log("updating to version " + details.version);
  chrome.runtime.reload();
});

// Fired when a context menu item is clicked.
chrome.contextMenus.onClicked.addListener(
  (info: chrome.contextMenus.OnClickData, tab?: chrome.tabs.Tab) => {
    // console.log("context menu clicked");
    // console.log(info);
    // console.log(tab);
    if (tab == null || tab.id == null || tab.id < 0) {
      return;
    }
    if (info.menuItemId === contextMenuId) {
      chrome.tabs.sendMessage(
        tab.id,
        {
          message: messageName,
        },
        (response: any) => { // Depending on your response, you may want to replace "any" with the actual type
          logError("tabs.sendMessage");
          if (response != null && response.message === successMessage) {
            // console.log("contextMenu operation succeeded.");
          }
        }
      );
    }
  }
);
