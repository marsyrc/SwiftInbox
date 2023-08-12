'use strict';

(function () {
    let _mailbox;
    let _customProps;

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            _mailbox = Office.context.mailbox;
            _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);

            loadItemProps(Office.context.mailbox.item);
        });
    });

    function loadItemProps(item) {
        // Write message property values to the task pane
        $('#item-id').text(item.itemId);
        $('#item-subject').text(item.subject);
        $('#item-internetMessageId').text(item.internetMessageId);
        $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");
    }

    // Callback function from loading custom properties.
    function customPropsCallback(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            console.error(asyncResult.status);
        }
        else {
            // Successfully loaded custom properties,
            // can get them from the asyncResult argument.
            _customProps = asyncResult.value;
            console.error("Successfully loaded custom properties")

            // Debug custom property implementation
            updateProperty("TestKey", "TestVal") ;
        }
    }

    // Set individual custom property.
    function updateProperty(name, value) {
        _customProps.set(name, value);
        // Save all custom properties to the mail item.
        _customProps.saveAsync(saveUpdateCallback);
    }

    // Get individual custom property.
    function getProperty() {
        if (_customProps) {
            const myPropValue = _customProps.get("TestKey");
            if (myPropValue) {
                console.log("Value of TestKey:", myPropValue);
            } else {
                console.log("TestKey not found.");
            }
        } else {
            console.log("Custom properties not loaded yet.");
        }
    }

    // Callback function from saving custom properties.
    function saveUpdateCallback(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            // Handle the failure.
            console.error("save callback error");
        } else {
            console.error("save callback succeeds");
            getProperty();
        }
    }
})();
