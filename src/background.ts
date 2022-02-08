(function () {
    'use strict';
    const contextMenuId = "teams-see-more-context-menu-id";
    const messageName = "teams-see-more-message";
    let teamsLoaded: { [key: number]: boolean; } = {};
    const isUrlMatched = (url?: string) => {
        if (url != null && url.indexOf("https://teams.microsoft.com/") >= 0) {
            return true;
        }
        return false;
    };

    // Fired when a tab is updated.
    chrome.tabs.onUpdated.addListener((tabId: number, changeInfo: chrome.tabs.TabChangeInfo, tab: chrome.tabs.Tab) => {
        if (tab == null) {
            return;
        }
        if (changeInfo.status !== "complete") {
            return;
        }
        if (isUrlMatched(tab.url)) {
            chrome.scripting.executeScript({
                target: { tabId },
                files: ["foreground.js"]
            }).then((value: chrome.scripting.InjectionResult[]) => {
                if (chrome.runtime.lastError) {
                    return;
                }
            }).catch((reason: any) => {
                if (chrome.runtime.lastError) {
                    return;
                }
            })
        }
    });

    const updateContextMenus = (force: boolean, tab?: chrome.tabs.Tab) => {
        if (tab == null || tab.id == null) {
            return;
        }
        const tabId = tab.id;
        const isUrlTeams = isUrlMatched(tab.url);
        const wasUrlTeams = teamsLoaded[tabId];
        if ((!force) && (wasUrlTeams === isUrlTeams)) {
            return;
        }
        if (isUrlTeams) {
            teamsLoaded[tabId] = true;
            // console.log(`Tab ${tabId} is now teams`);
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
            teamsLoaded[tabId] = false;
            // console.log(`Tab ${tabId} is no longer teams`);
            chrome.contextMenus.remove(contextMenuId, () => {
                if (chrome.runtime.lastError) {
                    return;
                }
            });
        }
    };

    // Fired when a tab is created.
    // Note that the tab's URL may not be set at the time this event fired, but you can listen to onUpdated events to be notified when a URL is set.
    chrome.tabs.onCreated.addListener((tab: chrome.tabs.Tab) => {
        updateContextMenus(true, tab);
    });

    // Fired when a tab is updated.
    chrome.tabs.onUpdated.addListener((tabId: number, changeInfo: chrome.tabs.TabChangeInfo, tab: chrome.tabs.Tab) => {
        updateContextMenus(false, tab);
    });

    // Fires when the active tab in a window changes.
    // Note that the tab's URL may not be set at the time this event fired, but you can listen to onUpdated events to be notified when a URL is set.
    chrome.tabs.onActivated.addListener((activeInfo: chrome.tabs.TabActiveInfo) => {
        chrome.tabs.get(activeInfo.tabId, (tab: chrome.tabs.Tab) => {
            updateContextMenus(true, tab);
        });
    });

    // Fired when a tab is closed.
    chrome.tabs.onRemoved.addListener((tabId: number, removeInfo: chrome.tabs.TabRemoveInfo) => {
        delete teamsLoaded[tabId];
    });

    // Fired when the currently focused window changes.
    // Will be chrome.windows.WINDOW_ID_NONE if all chrome windows have lost focus.
    chrome.windows.onFocusChanged.addListener((windowId: number) => {
        chrome.tabs.query({ active: true, currentWindow: true }, (tabs: chrome.tabs.Tab[]) => {
            const tab = tabs[0];
            updateContextMenus(true, tab);
        });
    });

    // Fired when a context menu item is clicked.
    chrome.contextMenus.onClicked.addListener((info: chrome.contextMenus.OnClickData, tab?: chrome.tabs.Tab) => {
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