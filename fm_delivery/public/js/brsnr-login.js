/* =========================================================
   BRSNR PORTAL LOGIN
   OTP Authentication
   Frappe / ERPNext 15
   ========================================================= */

(() => {
    "use strict";

    /* =========================================================
       CONFIG
       ========================================================= */

    const CONFIG = {
        SEND_OTP_METHOD: "fm_delivery.brsnr_api.send_login_otp",
        VERIFY_OTP_METHOD: "fm_delivery.brsnr_api.verify_login_otp",

        DASHBOARD_URL: "/brsnr-dashboard",

        OTP_LENGTH: 6,
        RESEND_SECONDS: 60,

        EMAIL_STORAGE_KEY: "brsnr_email",
        TOKEN_STORAGE_KEY: "brsnr_token",
        INCHARGE_STORAGE_KEY: "brsnr_incharge",
        ROLE_STORAGE_KEY: "brsnr_role",
        ASSIGNMENTS_STORAGE_KEY: "brsnr_assignments"
    };


    /* =========================================================
       DOM
       ========================================================= */

    const emailInput =
        document.getElementById("email") ||
        document.getElementById("login_email") ||
        document.querySelector('input[type="email"]');

    const otpInput =
        document.getElementById("otp") ||
        document.getElementById("login_otp") ||
        document.querySelector('input[inputmode="numeric"]');

    const sendBtn =
        document.getElementById("sendOtpBtn") ||
        document.getElementById("send-otp") ||
        document.getElementById("sendOTP") ||
        document.querySelector('[data-action="send-otp"]');

    const verifyBtn =
        document.getElementById("verifyOtpBtn") ||
        document.getElementById("verify-otp") ||
        document.getElementById("verifyOTP") ||
        document.querySelector('[data-action="verify-otp"]');

    const resendBtn =
        document.getElementById("resendOtpBtn") ||
        document.getElementById("resend-otp") ||
        document.querySelector('[data-action="resend-otp"]');

    const emailSection =
        document.getElementById("emailSection") ||
        document.getElementById("email-section") ||
        document.querySelector(".email-section");

    const otpSection =
        document.getElementById("otpSection") ||
        document.getElementById("otp-section") ||
        document.querySelector(".otp-section");

    const messageBox =
        document.getElementById("message") ||
        document.getElementById("errorMessage") ||
        document.getElementById("messageBox") ||
        document.querySelector(".message");

    const errorBox =
        document.getElementById("errorMessage") ||
        document.querySelector(".error-message");

    const successBox =
        document.getElementById("successMessage") ||
        document.querySelector(".success-message");

    let resendTimer = null;
    let resendRemaining = 0;


    /* =========================================================
       UTILITY
       ========================================================= */

    function getCsrfToken() {
        if (
            typeof frappe !== "undefined" &&
            frappe &&
            frappe.csrf_token
        ) {
            return frappe.csrf_token;
        }

        const meta = document.querySelector(
            'meta[name="csrf-token"]'
        );

        if (meta) {
            return meta.getAttribute("content") || "";
        }

        return "";
    }


    function isValidEmail(email) {
        return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
    }


    function normalizeEmail(email) {
        return String(email || "")
            .trim()
            .toLowerCase();
    }


    function getErrorMessage(data, fallback) {
        if (!data) {
            return fallback;
        }

        if (typeof data === "string") {
            return data;
        }

        if (typeof data.message === "string") {
            return data.message;
        }

        if (
            data.message &&
            typeof data.message.message === "string"
        ) {
            return data.message.message;
        }

        if (typeof data.exception === "string") {
            return data.exception;
        }

        return fallback;
    }


    /* =========================================================
       UI MESSAGE
       ========================================================= */

    function clearMessages() {
        if (messageBox) {
            messageBox.textContent = "";
            messageBox.style.display = "none";
            messageBox.classList.remove(
                "error",
                "success",
                "warning"
            );
        }

        if (errorBox) {
            errorBox.textContent = "";
            errorBox.style.display = "none";
        }

        if (successBox) {
            successBox.textContent = "";
            successBox.style.display = "none";
        }
    }


    function showError(message) {
        console.error("BRSNR:", message);

        if (errorBox) {
            errorBox.textContent = message;
            errorBox.style.display = "block";
            return;
        }

        if (messageBox) {
            messageBox.textContent = message;
            messageBox.style.display = "block";
            messageBox.classList.remove("success");
            messageBox.classList.add("error");
            return;
        }

        alert(message);
    }


    function showSuccess(message) {
        console.log("BRSNR:", message);

        if (successBox) {
            successBox.textContent = message;
            successBox.style.display = "block";
            return;
        }

        if (messageBox) {
            messageBox.textContent = message;
            messageBox.style.display = "block";
            messageBox.classList.remove("error");
            messageBox.classList.add("success");
        }
    }


    /* =========================================================
       API
       IMPORTANT:
       ALL AUTH REQUESTS ARE POST
       ========================================================= */

    async function apiPost(method, payload = {}) {

        const url = `/api/method/${method}`;

        const headers = {
            "Content-Type": "application/json",
            "Accept": "application/json"
        };

        const csrfToken = getCsrfToken();

        if (csrfToken) {
            headers["X-Frappe-CSRF-Token"] = csrfToken;
        }

        let response;

        try {

            response = await fetch(url, {
                method: "POST",
                credentials: "same-origin",
                headers: headers,
                body: JSON.stringify(payload)
            });

        } catch (networkError) {

            console.error(
                "Network error:",
                networkError
            );

            throw new Error(
                "Unable to connect to the server. Please try again."
            );
        }


        /* -----------------------------------------------------
           IMPORTANT:
           Do NOT directly use response.json().
           If Frappe returns HTML error page,
           response.json() gives:
           Unexpected token '<'
           ----------------------------------------------------- */

        const rawText = await response.text();

        let data = null;

        try {

            data = rawText
                ? JSON.parse(rawText)
                : null;

        } catch (parseError) {

            console.error(
                "Non-JSON server response:",
                rawText
            );

            throw new Error(
                `Server returned HTTP ${response.status}. ` +
                `Please check the Frappe error log.`
            );
        }


        if (!response.ok) {

            const message = getErrorMessage(
                data,
                `Request failed with HTTP ${response.status}`
            );

            throw new Error(message);
        }


        /*
         * Frappe usually returns:
         *
         * {
         *     "message": {
         *         ...
         *     }
         * }
         */

        return data?.message ?? data;
    }


    /* =========================================================
       LOADING STATE
       ========================================================= */

    function setButtonLoading(button, loading, loadingText) {

        if (!button) {
            return;
        }

        if (loading) {

            if (!button.dataset.originalText) {
                button.dataset.originalText =
                    button.textContent;
            }

            button.disabled = true;
            button.classList.add("loading");

            if (loadingText) {
                button.textContent = loadingText;
            }

        } else {

            button.disabled = false;
            button.classList.remove("loading");

            if (button.dataset.originalText) {
                button.textContent =
                    button.dataset.originalText;
            }
        }
    }


    /* =========================================================
       SHOW / HIDE SECTIONS
       ========================================================= */

    function showOTPSection() {

        if (emailSection) {
            emailSection.style.display = "none";
        }

        if (otpSection) {
            otpSection.style.display = "block";
        }

        if (otpInput) {
            otpInput.value = "";
            setTimeout(() => {
                otpInput.focus();
            }, 100);
        }
    }


    function showEmailSection() {

        if (emailSection) {
            emailSection.style.display = "block";
        }

        if (otpSection) {
            otpSection.style.display = "none";
        }

        if (emailInput) {
            setTimeout(() => {
                emailInput.focus();
            }, 100);
        }
    }


    /* =========================================================
       RESEND TIMER
       ========================================================= */

    function updateResendButton() {

        if (!resendBtn) {
            return;
        }

        if (resendRemaining > 0) {

            resendBtn.disabled = true;

            const original =
                resendBtn.dataset.originalText ||
                "Resend OTP";

            resendBtn.textContent =
                `${original} (${resendRemaining}s)`;

        } else {

            resendBtn.disabled = false;

            resendBtn.textContent =
                resendBtn.dataset.originalText ||
                "Resend OTP";
        }
    }


    function startResendTimer() {

        clearInterval(resendTimer);

        resendRemaining =
            CONFIG.RESEND_SECONDS;

        if (resendBtn) {

            if (!resendBtn.dataset.originalText) {
                resendBtn.dataset.originalText =
                    resendBtn.textContent ||
                    "Resend OTP";
            }
        }

        updateResendButton();

        resendTimer = setInterval(() => {

            resendRemaining--;

            updateResendButton();

            if (resendRemaining <= 0) {

                clearInterval(resendTimer);
                resendTimer = null;
                resendRemaining = 0;

                updateResendButton();
            }

        }, 1000);
    }


    /* =========================================================
       SEND OTP
       ========================================================= */

    async function sendOTP() {

        clearMessages();

        if (!emailInput) {

            showError(
                "Email input field was not found."
            );

            return;
        }

        const email =
            normalizeEmail(emailInput.value);


        /* Validation */

        if (!email) {

            showError(
                "Please enter your email address."
            );

            emailInput.focus();
            return;
        }


        if (!isValidEmail(email)) {

            showError(
                "Please enter a valid email address."
            );

            emailInput.focus();
            return;
        }


        /* Prevent resend during cooldown */

        if (resendRemaining > 0) {

            showError(
                `Please wait ${resendRemaining} seconds before requesting another OTP.`
            );

            return;
        }


        setButtonLoading(
            sendBtn,
            true,
            "Sending OTP..."
        );


        try {

            /*
             * IMPORTANT
             *
             * This is POST.
             *
             * OLD / WRONG:
             *
             * GET /api/method/...send_login_otp?email=...
             *
             * NEW / CORRECT:
             *
             * POST /api/method/...send_login_otp
             *
             * Body:
             * {
             *     "email": "..."
             * }
             */

            const result = await apiPost(
                CONFIG.SEND_OTP_METHOD,
                {
                    email: email
                }
            );


            console.log(
                "OTP response:",
                result
            );


            /*
             * Backend intentionally gives a generic
             * response for login-security purposes.
             */

            sessionStorage.setItem(
                CONFIG.EMAIL_STORAGE_KEY,
                email
            );


            showOTPSection();

            startResendTimer();

            showSuccess(
                "If the email is registered for BRSNR access, an OTP has been sent."
            );


        } catch (error) {

            console.error(
                "sendOTP error:",
                error
            );

            showError(
                error.message ||
                "Unable to send OTP. Please try again."
            );

        } finally {

            setButtonLoading(
                sendBtn,
                false
            );
        }
    }


    /* =========================================================
       VERIFY OTP
       ========================================================= */

    async function verifyOTP() {

        clearMessages();

        const email =
            normalizeEmail(
                emailInput?.value ||
                sessionStorage.getItem(
                    CONFIG.EMAIL_STORAGE_KEY
                )
            );

        const otp =
            String(
                otpInput?.value || ""
            ).replace(/\D/g, "");


        /* Email validation */

        if (!email) {

            showError(
                "Email address is missing."
            );

            showEmailSection();
            return;
        }


        /* OTP validation */

        if (!otp) {

            showError(
                "Please enter the OTP."
            );

            otpInput?.focus();
            return;
        }


        if (otp.length !== CONFIG.OTP_LENGTH) {

            showError(
                `Please enter the ${CONFIG.OTP_LENGTH}-digit OTP.`
            );

            otpInput?.focus();
            return;
        }


        setButtonLoading(
            verifyBtn,
            true,
            "Verifying..."
        );


        try {

            /*
             * IMPORTANT:
             *
             * verify_login_otp is also POST.
             */

            const result = await apiPost(
                CONFIG.VERIFY_OTP_METHOD,
                {
                    email: email,
                    otp: otp
                }
            );


            console.log(
                "OTP verification response:",
                result
            );


            if (!result) {

                throw new Error(
                    "Invalid server response."
                );
            }


            /*
             * Expected backend response can contain:
             *
             * {
             *     success: true,
             *     token: "...",
             *     email: "...",
             *     role: "...",
             *     incharge: "...",
             *     assignments: [...]
             * }
             */


            if (
                result.success === false
            ) {

                throw new Error(
                    result.message ||
                    "Invalid or expired OTP."
                );
            }


            const token =
                result.token ||
                result.session_token ||
                result.access_token;


            if (!token) {

                console.error(
                    "Verification response did not contain token:",
                    result
                );

                throw new Error(
                    "Login succeeded but no session token was received."
                );
            }


            /* -------------------------------------------------
               STORE AUTH SESSION
               ------------------------------------------------- */

            sessionStorage.setItem(
                CONFIG.TOKEN_STORAGE_KEY,
                token
            );


            sessionStorage.setItem(
                CONFIG.EMAIL_STORAGE_KEY,
                email
            );


            if (result.incharge) {

                sessionStorage.setItem(
                    CONFIG.INCHARGE_STORAGE_KEY,
                    typeof result.incharge === "string"
                        ? result.incharge
                        : JSON.stringify(result.incharge)
                );
            }


            if (result.role) {

                sessionStorage.setItem(
                    CONFIG.ROLE_STORAGE_KEY,
                    result.role
                );
            }


            if (result.assignments) {

                sessionStorage.setItem(
                    CONFIG.ASSIGNMENTS_STORAGE_KEY,
                    JSON.stringify(
                        result.assignments
                    )
                );
            }


            showSuccess(
                "Login successful. Redirecting..."
            );


            /*
             * Small delay so user sees success message.
             */

            setTimeout(() => {

                window.location.href =
                    CONFIG.DASHBOARD_URL;

            }, 300);


        } catch (error) {

            console.error(
                "verifyOTP error:",
                error
            );

            showError(
                error.message ||
                "Invalid or expired OTP."
            );


        } finally {

            setButtonLoading(
                verifyBtn,
                false
            );
        }
    }


    /* =========================================================
       RESEND OTP
       ========================================================= */

    async function resendOTP() {

        clearMessages();

        if (resendRemaining > 0) {

            showError(
                `Please wait ${resendRemaining} seconds.`
            );

            return;
        }

        await sendOTP();
    }


    /* =========================================================
       OTP INPUT
       ========================================================= */

    function setupOTPInput() {

        if (!otpInput) {
            return;
        }


        otpInput.setAttribute(
            "maxlength",
            String(CONFIG.OTP_LENGTH)
        );

        otpInput.setAttribute(
            "inputmode",
            "numeric"
        );

        otpInput.setAttribute(
            "autocomplete",
            "one-time-code"
        );


        otpInput.addEventListener(
            "input",
            () => {

                otpInput.value =
                    otpInput.value
                        .replace(/\D/g, "")
                        .slice(
                            0,
                            CONFIG.OTP_LENGTH
                        );
            }
        );


        otpInput.addEventListener(
            "keydown",
            event => {

                if (
                    event.key === "Enter"
                ) {

                    event.preventDefault();

                    verifyOTP();
                }
            }
        );
    }


    /* =========================================================
       EMAIL INPUT
       ========================================================= */

    function setupEmailInput() {

        if (!emailInput) {
            return;
        }


        emailInput.addEventListener(
            "input",
            () => {

                clearMessages();

                emailInput.value =
                    emailInput.value
                        .replace(/\s/g, "")
                        .toLowerCase();
            }
        );


        emailInput.addEventListener(
            "keydown",
            event => {

                if (
                    event.key === "Enter"
                ) {

                    event.preventDefault();

                    sendOTP();
                }
            }
        );
    }


    /* =========================================================
       BUTTON EVENTS
       ========================================================= */

    function setupEvents() {

        if (sendBtn) {

            sendBtn.addEventListener(
                "click",
                event => {

                    event.preventDefault();

                    sendOTP();
                }
            );
        }


        if (verifyBtn) {

            verifyBtn.addEventListener(
                "click",
                event => {

                    event.preventDefault();

                    verifyOTP();
                }
            );
        }


        if (resendBtn) {

            resendBtn.addEventListener(
                "click",
                event => {

                    event.preventDefault();

                    resendOTP();
                }
            );
        }


        /*
         * Support forms too.
         */

        const loginForm =
            document.getElementById("loginForm") ||
            document.querySelector(
                "form[data-brsnr-login]"
            );

        if (loginForm) {

            loginForm.addEventListener(
                "submit",
                event => {

                    event.preventDefault();

                    /*
                     * If OTP section is visible,
                     * verify OTP.
                     */

                    const otpVisible =
                        otpSection &&
                        getComputedStyle(
                            otpSection
                        ).display !== "none";

                    if (otpVisible) {

                        verifyOTP();

                    } else {

                        sendOTP();
                    }
                }
            );
        }
    }


    /* =========================================================
       EXISTING SESSION
       ========================================================= */

    function checkExistingSession() {

        const token =
            sessionStorage.getItem(
                CONFIG.TOKEN_STORAGE_KEY
            );

        if (!token) {
            return;
        }


        /*
         * Do NOT automatically redirect blindly.
         *
         * Dashboard/backend will validate token.
         *
         * If token expired, dashboard should redirect
         * back to login.
         */
    }


    /* =========================================================
       INITIALIZATION
       ========================================================= */

    function init() {

        console.log(
            "BRSNR Login initialized"
        );

        console.log(
            "OTP API:",
            CONFIG.SEND_OTP_METHOD
        );

        setupEmailInput();

        setupOTPInput();

        setupEvents();

        checkExistingSession();


        /*
         * Restore email if available.
         */

        const savedEmail =
            sessionStorage.getItem(
                CONFIG.EMAIL_STORAGE_KEY
            );

        if (
            savedEmail &&
            emailInput &&
            !emailInput.value
        ) {

            emailInput.value =
                savedEmail;
        }
    }


    /* =========================================================
       GLOBAL FUNCTIONS
       =========================================================
       Useful if your HTML uses onclick="sendOTP()"
       ========================================================= */

    window.sendOTP = sendOTP;
    window.verifyOTP = verifyOTP;
    window.resendOTP = resendOTP;


    /* =========================================================
       START
       ========================================================= */

    if (
        document.readyState === "loading"
    ) {

        document.addEventListener(
            "DOMContentLoaded",
            init
        );

    } else {

        init();
    }

})();