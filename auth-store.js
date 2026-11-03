/* =====================================================================
   Where the sign-in is kept.

   The site used to live entirely on one address, so supabase-js could
   keep your sign-in in localStorage and every page found it there.
   localStorage belongs to one origin, and subdomains are different
   origins: a session saved on civilinitiatives.org is invisible to
   learn.civilinitiatives.org. Signed in on the Launchpad, signed out
   in Level Up.

   A cookie can be scoped to a whole domain, so this keeps the session
   in one scoped to .civilinitiatives.org and every subdomain reads the
   same thing. supabase-js lets you hand it any storage object with
   getItem/setItem/removeItem, so nothing else has to change.

   Two things this has to be careful about:

   - A cookie holds about 4KB. A session (two tokens plus the user) can
     be larger, so a long value is split across several cookies and put
     back together on the way out.
   - A cookie for .civilinitiatives.org cannot be set from anywhere
     else. On the lab, which is served from github.io, this falls back
     to localStorage and behaves exactly as before. (github.io is on the
     Public Suffix List, so no shared cookie is possible there at all -
     that is a browser rule, not something to work around.)

   Anyone already signed in is carried across: the first time a value is
   asked for and the cookie is empty, it is taken from localStorage and
   written to the cookie. Nobody gets logged out by this change.

   Load this BEFORE creating the Supabase client, then:
     supabase.createClient(URL, KEY, {
       auth: { persistSession:true, autoRefreshToken:true,
               storage: window.HEHAuth.storage }
     });
   ===================================================================== */
(function () {
  var ROOT = "civilinitiatives.org";
  var host = location.hostname;
  var onDomain = (host === ROOT) || (host.length > ROOT.length &&
                  host.slice(-(ROOT.length + 1)) === "." + ROOT);

  /* Each cookie stays well under the 4KB limit. The value is
     percent-encoded first and the encoded text is what gets split, so a
     chunk boundary can never land in the middle of an escape. */
  var CHUNK = 3000;
  var MAXAGE = 60 * 60 * 24 * 400;   // the longest a browser will honour

  function jar() {
    var out = {}, all = document.cookie ? document.cookie.split("; ") : [];
    for (var i = 0; i < all.length; i++) {
      var eq = all[i].indexOf("=");
      if (eq > 0) out[all[i].slice(0, eq)] = all[i].slice(eq + 1);
    }
    return out;
  }

  function put(name, value) {
    document.cookie = name + "=" + value +
      "; domain=." + ROOT + "; path=/; max-age=" + MAXAGE +
      "; secure; samesite=lax";
  }

  function drop(name) {
    document.cookie = name + "=; domain=." + ROOT +
      "; path=/; max-age=0; secure; samesite=lax";
  }

  function clearKey(key) {
    var c = jar();
    if (c[key] !== undefined) drop(key);
    for (var k in c) if (k.indexOf(key + ".") === 0) drop(k);
  }

  var cookieStore = {
    getItem: function (key) {
      var c = jar(), encoded = null;
      if (c[key] !== undefined) {
        encoded = c[key];                       // short enough to be whole
      } else if (c[key + ".0"] !== undefined) {
        encoded = "";
        for (var i = 0; c[key + "." + i] !== undefined; i++) encoded += c[key + "." + i];
      }
      if (encoded === null) return null;
      try { return decodeURIComponent(encoded); } catch (e) { return null; }
    },
    setItem: function (key, value) {
      clearKey(key);
      var enc = encodeURIComponent(String(value));
      if (enc.length <= CHUNK) { put(key, enc); return; }
      for (var i = 0, n = 0; i < enc.length; i += CHUNK, n++) {
        put(key + "." + n, enc.substr(i, CHUNK));
      }
    },
    removeItem: function (key) { clearKey(key); }
  };

  var localStore = {
    getItem: function (key) { try { return localStorage.getItem(key); } catch (e) { return null; } },
    setItem: function (key, value) { try { localStorage.setItem(key, value); } catch (e) {} },
    removeItem: function (key) { try { localStorage.removeItem(key); } catch (e) {} }
  };

  /* On the domain: cookies, with localStorage read once as a fallback so
     an existing sign-in survives. Signing out has to clear both, or the
     old localStorage copy would bring the session back from the dead on
     the next page load. */
  var shared = {
    getItem: function (key) {
      var v = cookieStore.getItem(key);
      if (v !== null && v !== undefined) return v;
      var old = localStore.getItem(key);
      if (old !== null && old !== undefined) { cookieStore.setItem(key, old); return old; }
      return null;
    },
    setItem: function (key, value) {
      cookieStore.setItem(key, value);
      localStore.removeItem(key);      // one copy only, so they cannot disagree
    },
    removeItem: function (key) {
      cookieStore.removeItem(key);
      localStore.removeItem(key);
    }
  };

  window.HEHAuth = {
    onDomain: onDomain,
    storage: onDomain ? shared : localStore,
    /* For code that wants the token without going through supabase-js. */
    get: function (key) { return this.storage.getItem(key); },
    /* "Is anyone signed in?", without waiting for getSession() to answer.
       The key is sb-<project>-auth-token, so look for that shape in both
       places rather than hard-coding the project name. */
    hasSession: function () {
      var k, i;
      var c = jar();
      for (k in c) if (/^sb-.*-auth-token(\.\d+)?$/.test(k) && c[k]) return true;
      try {
        for (i = 0; i < localStorage.length; i++) {
          k = localStorage.key(i);
          if (k && /^sb-.*-auth-token$/.test(k) && localStorage.getItem(k)) return true;
        }
      } catch (e) {}
      return false;
    }
  };
})();
