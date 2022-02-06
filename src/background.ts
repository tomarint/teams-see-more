(function () {
    'use strict';
    const contextMenuId = "teams-see-more-context-menu-id";
    const messageName = "teams-see-more-message";
    const isUrlMatched = (url?: string) => {
        if (url != null && url.indexOf("https://teams.microsoft.com/") >= 0) {
            return true;
        }
        return false;
    };
    chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
        if (tab == null) {
            return;
        }
        if (isUrlMatched(tab.url)) {
            if (changeInfo.status === "complete") {
                chrome.scripting.executeScript({
                    target: { tabId },
                    files: ["foreground.js"]
                }).then(() => {
                    // console.log("foreground.js is inserted.")
                }).catch(err => {
                    // console.error(err);
                })
            }
        }
    });

    const updateContextMenus = (tab?: chrome.tabs.Tab) => {
        if (tab != null && isUrlMatched(tab.url)) {
            chrome.contextMenus.create({
                id: contextMenuId,
                title: "Teams See More", // chrome.i18n.getMessage("extName"),
                contexts: ["page"]
            }, () => {
                if (chrome.runtime.lastError) {
                    return;
                }
            });
        }
        else {
            chrome.contextMenus.remove(contextMenuId, () => {
                if (chrome.runtime.lastError) {
                    return;
                }
            });
        }
    };

    // Context menu on tab is created.
    chrome.tabs.onCreated.addListener((tab: chrome.tabs.Tab) => {
        updateContextMenus(tab);
    });

    // Context menu on tab is updated.
    chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
        updateContextMenus(tab);
    });

    // Context menu on tab is activated.
    chrome.tabs.onActivated.addListener((activeInfo) => {
        chrome.tabs.get(activeInfo.tabId, (tab) => {
            updateContextMenus(tab);
        });
    });

    chrome.windows.onFocusChanged.addListener((windowId: number) => {
        chrome.tabs.query({ active: true, currentWindow: true }, (tabs: chrome.tabs.Tab[]) => {
            const tab = tabs[0];
            updateContextMenus(tab);
        });
    });

    chrome.contextMenus.onClicked.addListener((info, tab) => {
        // console.log('context menu clicked');
        // console.log(info);
        // console.log(tab);
        if (tab == null || tab.id == null) {
            return;
        }
        if (info.menuItemId === contextMenuId) {
            chrome.tabs.sendMessage(
                tab.id,
                {
                    message: messageName
                },
                (response) => {
                    if (response != null && response.message === "success") {
                        // console.log("button#see_more succeeded.");
                    }
                }
            );
            return true;
        }
    });
})();