// Launchpad Suite for Android (GPL-3.0).
//
// The Suite in the app's own web view. What a web view can't do by itself
// the app does for it: picking files, saving what the apps export (to
// Downloads), printing / saving as PDF, and Google sign-in (Google won't
// sign in inside a web view, so that goes to the browser and comes back as
// hehsuite://login with the session).
package org.civilinitiatives.suite

import android.app.Activity
import android.content.ContentValues
import android.content.Intent
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.os.Environment
import android.print.PrintAttributes
import android.print.PrintManager
import android.provider.MediaStore
import android.util.Base64
import android.view.View
import android.webkit.*
import android.widget.Toast
import java.io.File

class MainActivity : Activity() {
    private val home = "https://suite.civilinitiatives.org/"
    private val launchpad = "https://civilinitiatives.org/"
    private lateinit var web: WebView
    private var fileCallback: ValueCallback<Array<Uri>>? = null
    private val pickFile = 7

    override fun onCreate(saved: Bundle?) {
        super.onCreate(saved)
        web = WebView(this)
        setContentView(web)
        val s = web.settings
        s.javaScriptEnabled = true
        s.domStorageEnabled = true
        s.databaseEnabled = true
        s.allowFileAccess = false
        s.mediaPlaybackRequiresUserGesture = true
        s.userAgentString = s.userAgentString + " LaunchpadApp/2.0 LaunchpadSuite/1.0 (Android)"
        CookieManager.getInstance().setAcceptCookie(true)
        CookieManager.getInstance().setAcceptThirdPartyCookies(web, true)
        web.addJavascriptInterface(Bridge(), "SuiteAndroid")
        web.webViewClient = Client()
        web.webChromeClient = Chrome()
        if (saved != null) web.restoreState(saved)
        else if (!handle(intent)) web.loadUrl(home)
    }

    override fun onSaveInstanceState(out: Bundle) { super.onSaveInstanceState(out); web.saveState(out) }
    override fun onNewIntent(i: Intent) { super.onNewIntent(i); handle(i) }

    /* hehsuite://login#at=..&rt=.. (or the same as extras) after Google sign-in */
    private fun handle(i: Intent?): Boolean {
        val d = i?.data ?: return false
        if (d.scheme != "hehsuite") return false
        var at = i.getStringExtra("at"); var rt = i.getStringExtra("rt")
        if (at == null || rt == null) {
            for (p in (d.fragment ?: "").split("&")) {
                val kv = p.split("=", limit = 2); if (kv.size != 2) continue
                if (kv[0] == "at") at = Uri.decode(kv[1]); if (kv[0] == "rt") rt = Uri.decode(kv[1])
            }
        }
        if (at.isNullOrEmpty() || rt.isNullOrEmpty()) return false
        web.loadUrl(launchpad + "?appsession=1#at=" + Uri.encode(at) + "&rt=" + Uri.encode(rt))
        return true
    }

    private fun ours(u: Uri): Boolean {
        val h = u.host ?: return false
        return h == "civilinitiatives.org" || h.endsWith(".civilinitiatives.org") || h == "hehs-live.duckdns.org" ||
            h.endsWith("login.microsoftonline.com") || h.endsWith("login.live.com") || h.endsWith("microsoft.com") || h.endsWith("msauth.net") || h.endsWith("msftauth.net")
    }

    private inner class Client : WebViewClient() {
        override fun shouldOverrideUrlLoading(v: WebView, r: WebResourceRequest): Boolean {
            val u = r.url
            if (u.host == "hehs-live.duckdns.org" && (u.path ?: "").startsWith("/auth/v1/authorize") && (u.getQueryParameter("provider") ?: "google") == "google") {
                open(Uri.parse(launchpad + "?applogin=google&ret=suiteandroid")); return true
            }
            if (u.scheme == "https" && ours(u)) return false
            if (u.scheme == "blob" || u.scheme == "data") return false
            open(u); return true
        }
        override fun onPageFinished(v: WebView, url: String) {
            // print() and PDF go to Android's printing, which can also save a PDF
            v.evaluateJavascript("window.print=function(){SuiteAndroid.print(document.title||'Launchpad Suite')};", null)
        }
    }

    private inner class Chrome : WebChromeClient() {
        override fun onShowFileChooser(view: WebView, cb: ValueCallback<Array<Uri>>, params: FileChooserParams): Boolean {
            fileCallback?.onReceiveValue(null); fileCallback = cb
            return try {
                val i = Intent(Intent.ACTION_GET_CONTENT).addCategory(Intent.CATEGORY_OPENABLE).setType("*/*")
                if (params.acceptTypes.any { it.startsWith("image") }) i.type = "image/*"
                startActivityForResult(Intent.createChooser(i, "Open"), pickFile); true
            } catch (e: Exception) { fileCallback = null; false }
        }
    }

    @Deprecated("Deprecated in Java")
    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        if (requestCode == pickFile) {
            fileCallback?.onReceiveValue(WebChromeClient.FileChooserParams.parseResult(resultCode, data)); fileCallback = null; return
        }
        super.onActivityResult(requestCode, resultCode, data)
    }

    @Deprecated("Deprecated in Java")
    override fun onBackPressed() { if (web.canGoBack()) web.goBack() else super.onBackPressed() }

    private fun open(u: Uri) { try { startActivity(Intent(Intent.ACTION_VIEW, u)) } catch (e: Exception) { } }

    /** What the Suite's pages can ask the app to do. */
    inner class Bridge {
        @JavascriptInterface fun platform(): String = "android"
        /* a file the apps made (a .docx, .xlsx, .pptx, .pdf...) goes to Downloads */
        @JavascriptInterface fun save(name: String, mime: String, b64: String) {
            val safe = name.replace(Regex("[\\\\/:*?\"<>|]"), " ").take(120).ifBlank { "file" }
            try {
                val bytes = Base64.decode(b64, Base64.DEFAULT)
                if (Build.VERSION.SDK_INT >= 29) {
                    val cv = ContentValues().apply {
                        put(MediaStore.Downloads.DISPLAY_NAME, safe); put(MediaStore.Downloads.MIME_TYPE, mime)
                        put(MediaStore.Downloads.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS + "/Launchpad Suite")
                    }
                    val uri = contentResolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, cv) ?: throw Exception("no room")
                    contentResolver.openOutputStream(uri)?.use { it.write(bytes) }
                } else {
                    @Suppress("DEPRECATION")
                    val dir = File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS), "Launchpad Suite")
                    dir.mkdirs(); File(dir, safe).writeBytes(bytes)
                }
                runOnUiThread { Toast.makeText(this@MainActivity, "Saved to Downloads: $safe", Toast.LENGTH_LONG).show() }
            } catch (e: Exception) {
                runOnUiThread { Toast.makeText(this@MainActivity, "Couldn't save $safe", Toast.LENGTH_LONG).show() }
            }
        }
        @JavascriptInterface fun print(title: String) {
            runOnUiThread {
                val pm = getSystemService(PRINT_SERVICE) as PrintManager
                pm.print(title.take(80), web.createPrintDocumentAdapter(title.take(80)), PrintAttributes.Builder().build())
            }
        }
    }
}
