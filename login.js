// dont forget to include jQuery code
// preferably with .noConflict() in order not to break the site scripts
if (window.location.indexOf("ebs.szgdjt.com") > -1) {
    // Lets login to Gmail
    jQuery("#usernameField").val("youremail@gmail.com");
    jQuery("#passwordField").val("superSecretPassowrd");
    //jQuery("#gaia_loginform").submit();
}