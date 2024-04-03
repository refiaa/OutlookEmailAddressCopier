// ==UserScript==
// @name         Office365 Email Address Clipboard Copy
// @namespace    O365ClipboardEmailCopy
// @version      v0.0.5
// @description  Copy email address to clipboard
// @author       Refiaa
// @match        https://outlook.office.com/mail/*
// @grant        none
// ==/UserScript==

class ClipboardUtility {
    static copyTextToClipboard(text) {
        if (navigator.clipboard && window.isSecureContext) {
            navigator.clipboard.writeText(text).then(() => {
                console.log('Text copied to clipboard successfully.');
            }).catch(err => {
                console.error('Failed to copy text to clipboard.', err);
            });
        } 
        
        else { // Fallback for older browsers
            const textarea = document.createElement('textarea');
            textarea.value = text;
            document.body.appendChild(textarea);
            textarea.select();
            document.execCommand('copy');
            document.body.removeChild(textarea);
        }
    }
}

class NotificationManager {
    static notificationContainer = null;

    static showNotification(message) {
        const appContainer = document.getElementById('appContainer');
        if (!appContainer) {
            console.error('App container not found.');
            return;
        }

        if (this.notificationContainer) {
            this.notificationContainer.remove();
        }

        this.notificationContainer = document.createElement('div');
        this.notificationContainer.className = 'custom-notification-container';
        this.notificationContainer.style = 'position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); z-index: 1000;';

        const notificationElement = document.createElement('div');
        notificationElement.className = 'notification notification--in notification--success';
        notificationElement.style = 'background-color: #0078d4; color: white; padding: 10px; border-radius: 4px; margin-bottom: 10px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.2); display: flex; justify-content: space-between; align-items: center;';

        const messageSpan = document.createElement('span');
        messageSpan.textContent = message;
        messageSpan.style = 'flex-grow: 1; margin-right: 16px;';

        const closeButton = document.createElement('button');
        closeButton.innerHTML = '&times;';
        closeButton.style = 'background: none; border: none; color: white; cursor: pointer; font-size: 20px; line-height: 1; padding: 0;';

        closeButton.onclick = () => this.notificationContainer.remove();

        notificationElement.appendChild(messageSpan);
        notificationElement.appendChild(closeButton);

        this.notificationContainer.appendChild(notificationElement);

        appContainer.appendChild(this.notificationContainer);

        setTimeout(() => {
            if (this.notificationContainer) {
                this.notificationContainer.remove();
                this.notificationContainer = null;
            }
        }, 3000);
    }
}

class TooltipManager {
    constructor() {
        this.tooltip = this.createTooltip();
    }

    createTooltip() {
        const tooltip = document.createElement('div');
        tooltip.className = 'spiketip-tooltip';
        tooltip.innerHTML = `Copy Address<div class="spiketip" style="position: absolute; width: 0; height: 0; border-left: 5px solid transparent; border-right: 5px solid transparent; border-top: 5px solid #555; top: 100%; left: 50%; transform: translateX(-50%);"></div>`;
        tooltip.style.cssText = 'position: absolute; padding: 4px 8px; background-color: #555; color: white; border-radius: 4px; font-size: 12px; display: none; z-index: 1001; text-align: center;';
        document.body.appendChild(tooltip);
        return tooltip;
    }

    showTooltip(element) {
        this.tooltip.style.display = 'block';
        this.updateTooltipPosition(element);
    }

    hideTooltip() {
        this.tooltip.style.display = 'none';
    }

    updateTooltipPosition(element) {
        const rect = element.getBoundingClientRect();
        const tooltipWidth = this.tooltip.offsetWidth;
        const tooltipHeight = this.tooltip.offsetHeight;

        let posX = rect.left + (rect.width / 2) - (tooltipWidth / 2);
        let posY = rect.top - tooltipHeight - 8;

        if (posX < 0) posX = 0;
        if (posX + tooltipWidth > window.innerWidth) {
            posX = window.innerWidth - tooltipWidth;
        }

        this.tooltip.style.left = `${posX}px`;
        this.tooltip.style.top = `${posY}px`;
    }
}

class EmailAddressCopier {
    constructor() {
        this.tooltipManager = new TooltipManager();
        this.attachEmailCopyListeners();
        this.observeDOMChanges();
    }

    attachEmailCopyListeners() {
        const emailElement = document.getElementById('mectrl_currentAccount_secondary');
        if (emailElement) {
            emailElement.style.cursor = 'pointer';
            emailElement.addEventListener('click', () => {
                ClipboardUtility.copyTextToClipboard(emailElement.innerText.trim());
                NotificationManager.showNotification('Email address copied to clipboard');
            });

            emailElement.addEventListener('mouseenter', () => this.tooltipManager.showTooltip(emailElement));
            emailElement.addEventListener('mouseleave', () => this.tooltipManager.hideTooltip());
        }
    }

    observeDOMChanges() {
        const observer = new MutationObserver(() => {
            this.attachEmailCopyListeners();
        });

        observer.observe(document.body, { childList: true, subtree: true });
    }
}

(function() {
    'use strict';
    new EmailAddressCopier();
})();