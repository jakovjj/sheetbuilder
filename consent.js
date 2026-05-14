(function () {
    var CONSENT_KEY = 'sb_analytics_consent';
    var CLARITY_ID = 'wqwc6qcx8i';

    function injectClarity() {
        (function (c, l, a, r, i, t, y) {
            c[a] = c[a] || function () { (c[a].q = c[a].q || []).push(arguments); };
            t = l.createElement(r); t.async = 1; t.src = 'https://www.clarity.ms/tag/' + i;
            y = l.getElementsByTagName(r)[0]; y.parentNode.insertBefore(t, y);
        })(window, document, 'clarity', 'script', CLARITY_ID);
    }

    function enableAnalytics() {
        if (typeof gtag === 'function') {
            gtag('consent', 'update', { analytics_storage: 'granted' });
        }
        injectClarity();
    }

    function hideBanner() {
        var el = document.getElementById('cookie-consent-banner');
        if (el) el.hidden = true;
    }

    function showBanner() {
        var el = document.getElementById('cookie-consent-banner');
        if (el) el.hidden = false;
    }

    document.addEventListener('click', function (e) {
        if (e.target.id === 'cookie-accept') {
            localStorage.setItem(CONSENT_KEY, 'granted');
            hideBanner();
            enableAnalytics();
        } else if (e.target.id === 'cookie-decline') {
            localStorage.setItem(CONSENT_KEY, 'denied');
            hideBanner();
        }
    });

    var stored = localStorage.getItem(CONSENT_KEY);
    if (stored === 'granted') {
        enableAnalytics();
    } else if (!stored) {
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', showBanner);
        } else {
            showBanner();
        }
    }
})();
