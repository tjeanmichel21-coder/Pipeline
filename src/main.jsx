import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";

class ErrorScreen extends React.Component {
  constructor(p) { super(p); this.state = { err: null }; }
  static getDerivedStateFromError(err) { return { err }; }
  render() {
    if (this.state.err)
      return (
        <div style={{ fontFamily: "sans-serif", padding: 40, maxWidth: 700, margin: "0 auto", color: "#1a1a1a" }}>
          <h2 style={{ color: "#c23934" }}>Pipeline hit an error</h2>
          <pre style={{ whiteSpace: "pre-wrap", background: "#fdf2f0", padding: 16, borderRadius: 8, fontSize: 13, overflowX: "auto" }}>
            {String(this.state.err?.message || this.state.err)}
          </pre>
          <p>Screenshot this message and send it to Claude to get it fixed. Refreshing may help if it was temporary.</p>
        </div>
      );
    return <App />;
  }
}

window.addEventListener("error", (e) => {
  const root = document.getElementById("root");
  if (root && !root.childNodes.length) {
    root.innerHTML = '<div style="font-family:sans-serif;padding:40px;color:#c23934"><h2>Pipeline failed to start</h2><pre style="white-space:pre-wrap;background:#fdf2f0;padding:16px;border-radius:8px">' +
      String(e.message || e.error || "Unknown startup error").replace(/</g, "&lt;") +
      "</pre><p style='color:#1a1a1a'>Screenshot this and send it to Claude.</p></div>";
  }
});

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <ErrorScreen />
  </React.StrictMode>
);
