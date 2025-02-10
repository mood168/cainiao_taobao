// Function to create notification container
function createNotificationContainer() {
    let container = document.getElementById('notification-container');
    if (!container) {
        container = document.createElement('div');
        container.id = 'notification-container';
        container.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 10000;
        `;
        document.body.appendChild(container);
    }
    return container;
}

// Function to create notification message element
function createNotificationMessage(message, type) {
    const notification = document.createElement('div');
    notification.style.cssText = `
        padding: 10px 20px;
        margin-bottom: 10px;
        border-radius: 4px;
        color: white;
        min-width: 200px;
        max-width: 400px;
        position: relative;
        animation: slideIn 0.5s ease-in-out;
    `;

    // Set background color based on type
    switch (type) {
        case 'success':
            notification.style.backgroundColor = '#4CAF50';
            break;
        case 'error':
            notification.style.backgroundColor = '#f44336';
            break;
        default:
            notification.style.backgroundColor = '#2196F3';
    }

    notification.textContent = message;
    return notification;
}

// Function to show notification with auto-remove for info type
function showNotification(message, type = 'info') {
    try {
        console.log(message);
        const container = createNotificationContainer();
        if (!container) return;

        const notification = createNotificationMessage(message, type);
        if (!notification) return;

        container.appendChild(notification);

        // Auto-remove info messages after 2 seconds
        if (type === 'info') {
            setTimeout(() => {
                if (notification && notification.parentNode) {
                    notification.remove();
                }
            }, 2000);
            return;
        }

        // For success messages, remove after 5 seconds
        if (type === 'success') {
            setTimeout(() => {
                if (notification && notification.parentNode) {
                    notification.remove();
                }
            }, 5000);
            return;
        }

        // Error messages stay until manually closed
        const closeButton = document.createElement('span');
        closeButton.textContent = '×';
        closeButton.style.cssText = `
            position: absolute;
            right: 10px;
            top: 50%;
            transform: translateY(-50%);
            cursor: pointer;
            font-weight: bold;
        `;
        closeButton.onclick = () => notification.remove();
        notification.appendChild(closeButton);
    } catch (error) {
        console.error('Error showing notification:', error);
    }
}

// Add CSS animation
const style = document.createElement('style');
style.textContent = `
    @keyframes slideIn {
        from {
            transform: translateX(100%);
            opacity: 0;
        }
        to {
            transform: translateX(0);
            opacity: 1;
        }
    }
`;
document.head.appendChild(style); 