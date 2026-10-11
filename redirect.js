import {
    broadcastResponseToMainFrame
} from "@azure/msal-browser/redirect-bridge";


broadcastResponseToMainFrame()
    .then(() => {
        console.log(
            "PowerNaturally authentication response completed."
        );
    })
    .catch((error) => {
        console.error(
            "PowerNaturally authentication redirect failed:",
            error
        );

        document.body.innerHTML =
            "<p>PowerNaturally authentication could not be completed.</p>";
    });