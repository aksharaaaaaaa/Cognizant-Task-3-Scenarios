/**
 * Function to display primary contact of account in quick view on case form.
 * @param {Object} executionContext - The execution context.
 */
function displayPrimaryContact(executionContext) {
    try {
        var formContext = executionContext.getFormContext();
        
        // Retrieve customer details.
        var customer = getCustomerDetails(formContext);
        if (customer) {
            handleCustomerType(customer,formContext)
        }
    } catch (error) {
        console.log("Error displaying primary contact: " + error.message);
        Xrm.Navigation.openAlertDialog("Error displaying primary contact: " + error.message);
    }
}

/**
 * Function to get customer details from form context.
 * @param {Object} formContext - The form context.
 * @returns {Object|null} The customer details if found, else null.
 */
function getCustomerDetails(formContext) {
    const customer = formContext.getAttribute("customerid").getValue();
    if (customer) {
        const [customerDetails] = customer;
        const {id: customerId, entityType: customerType} = customerDetails;
        return {customerId, customerType};
    } else {
        return null;
    }
}

/**
 * Function to handle if customer type is account or contact and make primary contact field visible/required.
 * @param {Object} customer - The customer details
 * @param {Object} formContext - The form context.
 */
function handleCustomerType(customer, formContext) {
    if (customer.customerType === "account") {
        if (customer.customerId) {
            retrieveAccountContact(customer.customerId, formContext);
        }
        formContext.getAttribute("primarycontactid").setRequiredLevel("required");
        formContext.getControl("primarycontactid").setVisible(true);
    } else if (customer.customerType === "contact")
    {
        formContext.getAttribute("primarycontactid").setRequiredLevel("none");
        formContext.getControl("primarycontactid").setVisible(false);
    }
}

/**
 * Function to lookup account's primary contact.
 * @param {string} accountId - The account ID.
 * @param {Object} formContext - The form context.
 */
function retrieveAccountContact(accountId, formContext) {
    Xrm.WebApi.retrieveRecord("account", accountId, "?$select=_primarycontactid_value").then(
        (result) => {
            var contactId = result["_primarycontactid_value"];
            if (contactId) {
                setPrimaryContact(contactId, result, formContext);
            }
        },
        (error) => {
            console.log("Error retrieving account: " + error.message);
            Xrm.Navigation.openAlertDialog("Error retrieving account: " + error.message);
        }
    )
}

/**
 * Function to set primary contact on case form.
 * @param {string} contactId - The contact ID.
 * @param {Object} result - The result from the API call.
 * @param {Object} formContext - The form context.
 */
function setPrimaryContact(contactId, result, formContext) {
    var contactReference = [{
        id: contactId,
        entityType: "contact",
        name: result["_primarycontactid_value@OData.Community.Display.V1.FormattedValue"]
    }];
    if (formContext.getAttribute("primarycontactid").getValue() === null){
        formContext.getAttribute("primarycontactid").setValue(contactReference);
    }
}
 
/**
 * Function to set visibility of quick view fields based on if they are null.
 * @param {Object} executionContext - The execution context.
 */
function setQuickViewFieldsVisibility(executionContext) {
    try {
        var formContext = executionContext.getFormContext();
        var quickViewControl = formContext.ui.quickForms.get("PrimaryContactQV"); 
        if (quickViewControl.isLoaded()) {
            var email = quickViewControl.getControl("emailaddress1").getAttribute().getValue(); 
            var mobilePhone = quickViewControl.getControl("mobilephone").getAttribute().getValue();

            quickViewControl.getControl("emailaddress1").setVisible(!!email); 
            quickViewControl.getControl("mobilephone").setVisible(!!mobilePhone);
        } 
    } catch (error) {
        console.log("Error setting quick view visibility: " + error.message);
        Xrm.Navigation.openAlertDialog("Error setting quick view visibility: " + error.message);
    }
}

