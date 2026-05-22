// Button 1 handler - Отправить заказ
function onButton1Click() {
    const message = "✓ Заказ отправлен в ERP систему";
    showNotification(message);
}

// Button 2 handler - Отправить счет
function onButton2Click() {
    const message = "✓ Счет отправлен в ERP систему";
    showNotification(message);
}

// Button 3 handler - Отправить отчет
function onButton3Click() {
    const message = "✓ Отчет отправлен в ERP систему";
    showNotification(message);
}

// Function to display notification to user
function showNotification(message) {
    Office.onReady(function() {
        // Create a simple notification element
        const notification = document.createElement('div');
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background-color: #107c10;
            color: white;
            padding: 16px 24px;
            border-radius: 4px;
            font-size: 14px;
            z-index: 10000;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
            animation: slideIn 0.3s ease-out;
            font-family: Segoe UI, sans-serif;
        `;
        
        notification.textContent = message;
        document.body.appendChild(notification);
        
        // Add animation
        const style = document.createElement('style');
        style.textContent = `
            @keyframes slideIn {
                from {
                    transform: translateX(400px);
                    opacity: 0;
                }
                to {
                    transform: translateX(0);
                    opacity: 1;
                }
            }
            @keyframes slideOut {
                from {
                    transform: translateX(0);
                    opacity: 1;
                }
                to {
                    transform: translateX(400px);
                    opacity: 0;
                }
            }
        `;
        document.head.appendChild(style);
        
        // Remove notification after 3 seconds
        setTimeout(() => {
            notification.style.animation = 'slideOut 0.3s ease-out forwards';
            setTimeout(() => {
                document.body.removeChild(notification);
            }, 300);
        }, 3000);
    });
}

// Escape HTML to prevent injection
function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, m => map[m]);
}