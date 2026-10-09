frappe.ready(function () {
  "use strict";

  /* =====================================================================
     BRSNR Command Center - dashboard logic (compatible with fm_delivery.brsnr_api)
     Backend contract used (and ONLY this):
       get_portal_session   -> identity + hub_options
       get_dashboard_data   -> summary, records, records_total, trend,
                               hub_performance, risk, loss_register, hub_options
       logout_portal        -> revoke session token
     Auth: session token from sessionStorage, sent in the X-BRSNR-Token header
     (never in the URL, never logged, never in localStorage).
     Every number comes from the API. Nothing is hardcoded.
  ===================================================================== */

  var API = "fm_delivery.brsnr_api.";
  var MAX_PAGE_LIMIT = 2000;
  var RISK_LIMIT = 100;
  var LOSS_REGISTER_LIMIT = 1000;
  var MAX_REPORT_ROWS = 50000;

  var SORTABLE = {
    aging: 1, total_price: 1, shipment_id: 1, source: 1, current_hub: 1,
    brsnr_status: 1, clearance_status: 1, days_remaining: 1,
    loss_amount: 1, due_month: 1
  };

  var byId = function (id) {
    return document.getElementById(id);
  };

  var qs = function (s, r) {
    return (r || document).querySelector(s);
  };

  var qsa = function (s, r) {
    return Array.prototype.slice.call(
      (r || document).querySelectorAll(s)
    );
  };

  /* =============================== UTILS =============================== */

  var U = {
    esc: function (v) {
      if (v === null || v === undefined) return "";

      return String(v).replace(/[&<>"']/g, function (c) {
        return {
          "&": "&amp;",
          "<": "&lt;",
          ">": "&gt;",
          '"': "&quot;",
          "'": "&#39;"
        }[c];
      });
    },

    num: function (v) {
      return Number(v || 0).toLocaleString("en-IN");
    },

    dec: function (v, d) {
      return Number(v || 0).toLocaleString("en-IN", {
        minimumFractionDigits: d,
        maximumFractionDigits: d
      });
    },

    money: function (v) {
      return "₹" + Number(v || 0).toLocaleString("en-IN", {
        maximumFractionDigits: 2
      });
    },

    moneyOrDash: function (v) {
      return v === null || v === undefined || v === ""
        ? "—"
        : U.money(v);
    },

    moneyShort: function (v) {
      v = Number(v || 0);

      if (Math.abs(v) >= 1e7) {
        return "₹" + (v / 1e7).toFixed(2) + " Cr";
      }

      if (Math.abs(v) >= 1e5) {
        return "₹" + (v / 1e5).toFixed(2) + " L";
      }

      return U.money(v);
    },

    parse: function (v) {
      if (!v) return null;

      var s = String(v).replace(" ", "T");

      var d = new Date(
        s.length === 10
          ? s + "T00:00:00"
          : s
      );

      return isNaN(d.getTime()) ? null : d;
    },

    date: function (v) {
      var d = U.parse(v);

      return d
        ? d.toLocaleDateString("en-IN", {
            day: "2-digit",
            month: "short",
            year: "numeric"
          })
        : "—";
    },

    shortDate: function (v) {
      var d = U.parse(v);

      return d
        ? d.toLocaleDateString("en-IN", {
            day: "2-digit",
            month: "short"
          })
        : "";
    },

    month: function (v) {
      var d = U.parse(v);

      return d
        ? d.toLocaleDateString("en-IN", {
            month: "short",
            year: "numeric"
          })
        : "—";
    },

    dtime: function (v) {
      var d = U.parse(v);

      return d
        ? d.toLocaleString("en-IN", {
            day: "2-digit",
            month: "short",
            year: "numeric",
            hour: "2-digit",
            minute: "2-digit"
          })
        : "—";
    },

    iso: function (d) {
      return (
        d.getFullYear() +
        "-" +
        String(d.getMonth() + 1).padStart(2, "0") +
        "-" +
        String(d.getDate()).padStart(2, "0")
      );
    },

    val: function (v) {
      return v === null || v === undefined || v === ""
        ? "—"
        : U.esc(v);
    },

    debounce: function (fn, ms) {
      var t;

      return function () {
        var a = arguments;
        var c = this;

        clearTimeout(t);

        t = setTimeout(function () {
          fn.apply(c, a);
        }, ms);
      };
    },

    serverMsg: function (data) {
      try {
        if (data && data._server_messages) {
          var m = JSON.parse(
            JSON.parse(data._server_messages)[0]
          ).message;

          var t = document.createElement("div");
          t.innerHTML = m;

          return t.textContent;
        }
      } catch (e) {
        /* ignore */
      }

      return "";
    },

    csvCell: function (v) {
      if (v === null || v === undefined) return "";

      if (typeof v === "number") {
        return String(v);
      }

      var s = String(v);

      if (/^[=+\-@\t\r]/.test(s)) {
        s = "'" + s;
      }

      return /[",\n\r]/.test(s)
        ? '"' + s.replace(/"/g, '""') + '"'
        : s;
    }
  };

  /* =============================== AUTH =============================== */

  var Auth = {
    token: sessionStorage.getItem("brsnr_token"),

    clear: function () {
      [
        "brsnr_token",
        "brsnr_role",
        "brsnr_email",
        "brsnr_incharge",
        "brsnr_assignments"
      ].forEach(function (k) {
        sessionStorage.removeItem(k);
      });

      Auth.token = null;
    },

    toLogin: function () {
      window.location.href = "/brsnr-login";
    },

    expire: function () {
      Auth.clear();
      Auth.toLogin();
    },

    logout: function () {
      return Api.post("logout_portal", {})
        .catch(function () {})
        .then(function () {
          Auth.clear();
          Auth.toLogin();
        });
    }
  };

  if (!Auth.token) {
    Auth.toLogin();
    return;
  }

  /* =============================== API =============================== */

  function ApiError(message, status) {
    this.message = message;
    this.status = status;
  }

  ApiError.prototype = Object.create(Error.prototype);

  var Api = {
    post: function (method, args) {
      var body = new URLSearchParams();

      Object.keys(args || {}).forEach(function (k) {
        var v = args[k];

        if (
          v !== undefined &&
          v !== null &&
          v !== ""
        ) {
          body.set(k, v);
        }
      });

      return fetch(
        "/api/method/" + API + method,
        {
          method: "POST",
          credentials: "same-origin",
          body: body,

          headers: {
            "X-Frappe-CSRF-Token":
              window.csrf_token ||
              (window.frappe && frappe.csrf_token) ||
              "",

            "X-BRSNR-Token":
              Auth.token || "",

            "Accept":
              "application/json"
          }
        }
      )
        .then(function (res) {
          return res
            .json()
            .catch(function () {
              return {};
            })
            .then(function (data) {
              if (res.status === 401) {
                Auth.expire();

                throw new ApiError(
                  "Your session has expired. Please sign in again.",
                  401
                );
              }

              if (res.status === 403) {
                throw new ApiError(
                  U.serverMsg(data) ||
                  "You are not authorized for this data.",
                  403
                );
              }

              if (
                data.exc_type === "RateLimitExceededError" ||
                res.status === 429
              ) {
                throw new ApiError(
                  "Too many requests. Please wait a moment.",
                  429
                );
              }

              if (!res.ok) {
                throw new ApiError(
                  U.serverMsg(data) ||
                  ("Request failed (" + res.status + ")."),
                  res.status
                );
              }

              return data.message;
            });
        })
        .catch(function (e) {
          if (e instanceof ApiError) {
            throw e;
          }

          throw new ApiError(
            "Network error. Check your connection and retry.",
            0
          );
        });
    }
  };

  /* =============================== STATE =============================== */

  var SECTIONS = {
    overview: "Overview",
    shipments: "Live Shipments",
    pending: "Pending",
    "new": "New",
    cleared: "Cleared",
    risk: "Risk Monitor",
    analytics: "Analytics",
    hubperf: "Hub Performance",
    history: "History",
    reports: "Reports",
    profile: "Profile"
  };

  var NO_FILTERBAR = {
    history: 1,
    reports: 1,
    profile: 1
  };

  var S = {
    section: "overview",

    session: null,
    data: null,
    dataKey: "",
    latest: null,
    hubSig: "",

    loaded: {},
    seq: {},
    rec: {},
    safeCount: null,

    f: {
      from: "",
      to: "",
      hub: "",
      status: "",
      clearance: "",
      search: "",
      range: ""
    },

    ship: {
      page: 1,
      size: 50,
      sort: "aging",
      dir: "desc",
      total: 0
    },

    risk: {
      bucket: "Critical",
      page: 1,
      size: 50,
      total: 0
    },

    pendSort: "aging:desc",
    dates: []
  };

  function guard(key) {
    var id =
      (S.seq[key] = (S.seq[key] || 0) + 1);

    return function () {
      return S.seq[key] === id;
    };
  }

  function invalidate() {
    S.loaded = {};
  }

  /* =============================== UI HELPERS =============================== */

  function toast(msg, tone) {
    var el = document.createElement("div");

    el.className = "brsnr-toast";
    el.setAttribute(
      "data-tone",
      tone || "info"
    );

    el.textContent = msg;

    byId("brsnr-toasts").appendChild(el);

    setTimeout(function () {
      el.remove();
    }, 4200);
  }

  function setLive(state, text) {
    var chip = byId("brsnr-live");

    chip.setAttribute(
      "data-state",
      state
    );

    byId("brsnr-live-text").textContent =
      text ||
      {
        live: "Live",
        loading: "Updating",
        error: "Error"
      }[state];

    if (state === "live") {
      byId("brsnr-banner").hidden = true;

      byId(
        "brsnr-last-refresh"
      ).textContent =
        new Date().toLocaleTimeString(
          "en-IN",
          {
            hour: "2-digit",
            minute: "2-digit"
          }
        );
    }

    qsa("#brsnr-refresh-top").forEach(
      function (b) {
        b.classList.toggle(
          "brsnr-spin",
          state === "loading"
        );
      }
    );
  }

  function fail(e, tbody, cols) {
    if (e && e.status === 401) {
      return;
    }

    setLive("error");

    var msg =
      (e && e.message) ||
      "Something went wrong.";

    byId(
      "brsnr-banner-text"
    ).textContent = msg;

    byId("brsnr-banner").hidden = false;

    if (tbody) {
      tbody.innerHTML = R.emptyRow(
        cols || 6,
        "Could not load data",
        msg
      );
    }
  }

  /* =============================== THEME =============================== */

  var Theme = {
    get: function () {
      return (
        document.documentElement.getAttribute(
          "data-brsnr-theme"
        ) || "light"
      );
    },

    set: function (t) {
      document.documentElement.setAttribute(
        "data-brsnr-theme",
        t
      );

      try {
        localStorage.setItem(
          "brsnr_theme",
          t
        );
      } catch (e) {
        /* private mode */
      }

      byId("brsnr-theme").setAttribute(
        "aria-checked",
        String(t === "dark")
      );

      Charts.retheme();
    },

    toggle: function () {
      Theme.set(
        Theme.get() === "dark"
          ? "light"
          : "dark"
      );
    }
  };

  /* =============================== CHARTS =============================== */

  var Charts = {
    reg: {},

    css: function (n) {
      return getComputedStyle(
        document.documentElement
      )
        .getPropertyValue(n)
        .trim();
    },

    tone: function (t) {
      var m = {
        ok: "--brsnr-ok",
        info: "--brsnr-info",
        warn: "--brsnr-warn",
        crit: "--brsnr-crit",
        loss: "--brsnr-loss",
        accent: "--brsnr-accent",
        muted: "--brsnr-faint"
      };

      return Charts.css(
        m[t] || m.accent
      );
    },

    alpha: function (c, a) {
      var m =
        /^#([0-9a-f]{6})$/i.exec(c);

      if (!m) return c;

      var n = parseInt(m[1], 16);

      return (
        "rgba(" +
        (n >> 16) +
        "," +
        ((n >> 8) & 255) +
        "," +
        (n & 255) +
        "," +
        a +
        ")"
      );
    },

    clearOverlay: function (canvas) {
      qsa(
        ".brsnr-skel-chart,.brsnr-chart-empty",
        canvas.parentElement
      ).forEach(function (n) {
        n.remove();
      });
    },

    loading: function (id) {
      var c = byId(id);

      if (!c) return;

      Charts.clearOverlay(c);

      var s =
        document.createElement("span");

      s.className =
        "brsnr-skel brsnr-skel-chart";

      c.parentElement.appendChild(s);
    },

    empty: function (c, msg) {
      Charts.clearOverlay(c);

      var d =
        document.createElement("div");

      d.className =
        "brsnr-chart-empty";

      d.innerHTML =
        "<strong>No data</strong>" +
        "<span>" +
        U.esc(
          msg ||
          "Nothing to chart for this selection."
        ) +
        "</span>";

      c.parentElement.appendChild(d);
    },

    styleDatasets: function (
      type,
      datasets
    ) {
      datasets.forEach(function (ds) {
        if (type === "doughnut") {
          ds.backgroundColor =
            (ds.brsnrTones || [])
              .map(Charts.tone);

          ds.borderColor =
            Charts.css(
              "--brsnr-surface"
            );

          ds.borderWidth = 2;
        } else {
          var c =
            Charts.tone(
              ds.brsnrTone
            );

          if (type === "line") {
            ds.borderColor = c;
            ds.backgroundColor =
              Charts.alpha(c, 0.14);

            ds.tension = 0.35;
            ds.pointRadius = 2;
            ds.pointHoverRadius = 4;
            ds.borderWidth = 2;
            ds.spanGaps = true;
          } else {
            ds.backgroundColor =
              Charts.alpha(c, 0.85);

            ds.borderColor = c;
            ds.borderRadius = 4;
            ds.maxBarThickness = 38;
          }
        }
      });
    },

    options: function (
      type,
      nSeries,
      o
    ) {
      var text =
        Charts.css("--brsnr-muted");

      var grid =
        Charts.css(
          "--brsnr-grid-line"
        );

      var opt = {
        responsive: true,
        maintainAspectRatio: false,

        animation: {
          duration: 550
        },

        interaction: {
          mode:
            type === "doughnut"
              ? "nearest"
              : "index",
          intersect: false
        },

        plugins: {
          legend: {
            display:
              type === "doughnut" ||
              nSeries > 1,

            position:
              type === "doughnut"
                ? "bottom"
                : "top",

            labels: {
              color: text,
              usePointStyle: true,
              boxWidth: 8,
              padding: 14,
              font: {
                size: 11
              }
            }
          },

          tooltip: {
            backgroundColor:
              Charts.css(
                "--brsnr-surface"
              ),

            titleColor:
              Charts.css(
                "--brsnr-text"
              ),

            bodyColor: text,

            borderColor:
              Charts.css(
                "--brsnr-border"
              ),

            borderWidth: 1,
            padding: 10,
            cornerRadius: 8,
            boxPadding: 4,

            callbacks:
              o.tooltip || {}
          }
        }
      };

      if (type === "doughnut") {
        opt.cutout = "66%";
        return opt;
      }

      var cat = {
        grid: {
          display: false
        },

        ticks: {
          color: text,
          maxRotation: 0,
          autoSkip: true,

          font: {
            size: 11
          }
        },

        border: {
          color: grid
        }
      };

      var val = {
        beginAtZero: true,

        grid: {
          color: grid
        },

        border: {
          display: false
        },

        ticks: {
          color: text,

          font: {
            size: 11
          },

          callback:
            o.fmt ||
            function (v) {
              return U.num(v);
            }
        }
      };

      if (o.max) {
        val.max = o.max;
      }

      if (o.stacked) {
        cat.stacked = true;
        val.stacked = true;
      }

      if (o.horizontal) {
        opt.indexAxis = "y";
        opt.scales = {
          x: val,
          y: cat
        };
      } else {
        opt.scales = {
          x: cat,
          y: val
        };
      }

      return opt;
    },

    draw: function (
      id,
      type,
      labels,
      datasets,
      o
    ) {
      var canvas = byId(id);

      if (!canvas) return;

      o = o || {};

      if (typeof Chart === "undefined") {
        Charts.empty(
          canvas,
          "Chart library failed to load."
        );
        return;
      }

      var has =
        labels &&
        labels.length &&
        datasets.some(function (d) {
          return d.data.some(
            function (v) {
              return Number(v) > 0;
            }
          );
        });

      if (!has) {
        if (Charts.reg[id]) {
          Charts.reg[id].destroy();
          delete Charts.reg[id];
        }

        Charts.empty(
          canvas,
          o.emptyMsg
        );

        return;
      }

      Charts.clearOverlay(canvas);

      Charts.styleDatasets(
        type,
        datasets
      );

      var existing =
        Charts.reg[id];

      if (
        existing &&
        existing.config.type === type
      ) {
        existing.data.labels =
          labels;

        existing.data.datasets =
          datasets;

        existing.brsnr = {
          type: type,
          o: o
        };

        existing.options =
          Charts.options(
            type,
            datasets.length,
            o
          );

        existing.update();

        return;
      }

      if (existing) {
        existing.destroy();
      }

      var ch = new Chart(
        canvas,
        {
          type: type,

          data: {
            labels: labels,
            datasets: datasets
          },

          options:
            Charts.options(
              type,
              datasets.length,
              o
            )
        }
      );

      ch.brsnr = {
        type: type,
        o: o
      };

      Charts.reg[id] = ch;
    },

    na: function (id, msg) {
      Charts.draw(
        id,
        "bar",
        [],
        [],
        {
          emptyMsg:
            msg ||
            "Not provided by the dashboard API."
        }
      );
    },

    retheme: function () {
      Object.keys(
        Charts.reg
      ).forEach(function (id) {
        var ch =
          Charts.reg[id];

        if (!ch) return;

        Charts.styleDatasets(
          ch.brsnr.type,
          ch.data.datasets
        );

        ch.options =
          Charts.options(
            ch.brsnr.type,
            ch.data.datasets.length,
            ch.brsnr.o
          );

        ch.update("none");
      });
    }
  };

  /* =============================== RENDER HELPERS =============================== */

  var R = {
    badge: function (
      text,
      tone
    ) {
      return (
        '<span class="brsnr-badge" data-tone="' +
        tone +
        '">' +
        U.esc(text) +
        "</span>"
      );
    },

    statusTone: function (s) {
      return {
        New: "info",
        Pending: "warn",
        Cleared: "ok"
      }[s] || "muted";
    },

    clearTone: function (s) {
      return {
        Pending: "info",
        "Due Soon": "warn",
        Cleared: "ok",
        Loss: "loss"
      }[s] || "muted";
    },

    riskTone: function (s) {
      return {
        Safe: "ok",
        Watch: "warn",
        Critical: "crit",
        Loss: "loss"
      }[s] || "muted";
    },

    status: function (s) {
      return R.badge(
        s || "—",
        R.statusTone(s)
      );
    },

    clearance: function (s) {
      return R.badge(
        s === "Loss"
          ? "Past Due — Loss"
          : s || "—",
        R.clearTone(s)
      );
    },

    risk: function (s) {
      return R.badge(
        s || "—",
        R.riskTone(s)
      );
    },

    kpi: function (o) {
      var tag =
        o.goto
          ? "button"
          : "div";

      return (
        "<" +
        tag +
        ' class="brsnr-kpi"' +
        (o.tone
          ? ' data-tone="' +
            o.tone +
            '"'
          : "") +
        (o.goto
          ? ' data-goto="' +
            o.goto +
            '"' +
            (o.bucket
              ? ' data-bucket="' +
                o.bucket +
                '"'
              : "")
          : "") +
        (o.title
          ? ' title="' +
            U.esc(o.title) +
            '"'
          : "") +
        ">" +
        '<div class="brsnr-kpi-label">' +
        U.esc(o.label) +
        "</div>" +
        '<div class="brsnr-kpi-value">' +
        o.value +
        "</div>" +
        (o.sub
          ? '<div class="brsnr-kpi-sub">' +
            U.esc(o.sub) +
            "</div>"
          : "") +
        "</" +
        tag +
        ">"
      );
    },

    kpiNA: function (label) {
      return R.kpi({
        label: label,
        value: "—",
        sub: "Not provided by the dashboard API"
      });
    },

    skelKpis: function (
      el,
      n
    ) {
      el.innerHTML =
        new Array(n + 1).join(
          '<span class="brsnr-skel brsnr-skel-kpi"></span>'
        );
    },

    skelRows: function (
      tbody,
      cols,
      n
    ) {
      tbody.innerHTML =
        new Array(
          (n || 8) + 1
        ).join(
          '<tr><td colspan="' +
            cols +
            '"><span class="brsnr-skel brsnr-skel-row" style="margin:2px 0"></span></td></tr>'
        );
    },

    emptyRow: function (
      cols,
      title,
      msg
    ) {
      return (
        '<tr><td colspan="' +
        cols +
        '"><div class="brsnr-empty"><strong>' +
        U.esc(title) +
        "</strong>" +
        U.esc(msg || "") +
        "</div></td></tr>"
      );
    },

    bar: function (
      pct,
      tone
    ) {
      return (
        '<div class="brsnr-progress"' +
        (tone
          ? ' data-tone="' +
            tone +
            '"'
          : "") +
        '><i style="width:' +
        Math.max(
          0,
          Math.min(
            100,
            pct
          )
        ) +
        '%"></i></div>'
      );
    },

    clearBarTone: function (
      p
    ) {
      return p >= 70
        ? ""
        : p >= 40
        ? "warn"
        : "crit";
    },

    idCell: function (r) {
      return (
        '<span class="brsnr-id">' +
        U.esc(r.shipment_id) +
        "</span>"
      );
    },

    days: function (r) {
      return r.days_remaining ===
        null ||
        r.days_remaining ===
          undefined
        ? "—"
        : U.num(
            r.days_remaining
          );
    },

    dl: function (pairs) {
      return (
        '<dl class="brsnr-dl">' +
        pairs
          .map(function (p) {
            return (
              "<div><dt>" +
              U.esc(p[0]) +
              "</dt><dd>" +
              (p[2]
                ? p[1]
                : U.val(p[1])) +
              "</dd></div>"
            );
          })
          .join("") +
        "</dl>"
      );
    }
  };

  /* Row identity */
  function rk(r) {
    return (
      r.name ||
      String(r.source) +
        "|" +
        String(r.shipment_id)
    );
  }

  function reg(rows) {
    (rows || []).forEach(
      function (r) {
        S.rec[rk(r)] =
          Object.assign(
            S.rec[rk(r)] || {},
            r
          );
      }
    );
  }

  function rowOpen(r) {
    return (
      '<tr data-row="' +
      U.esc(rk(r)) +
      '">'
    );
  }

  /* =============================== TABLE ROWS =============================== */

  function shipRow(r) {
    return (
      rowOpen(r) +
      "<td>" +
      R.idCell(r) +
      "</td><td>" +
      U.esc(r.source) +
      "</td><td>" +
      U.val(r.current_hub) +
      "</td><td>" +
      U.val(r.partner) +
      "</td><td>" +
      R.status(r.brsnr_status) +
      "</td><td>" +
      R.clearance(
        r.clearance_status
      ) +
      "</td><td>" +
      R.risk(r.risk_level) +
      '</td><td class="brsnr-num">' +
      U.num(r.aging) +
      '</td><td class="brsnr-num">' +
      R.days(r) +
      '</td><td class="brsnr-num">' +
      U.money(r.total_price) +
      "</td><td>" +
      U.date(r.brsnr_date) +
      '</td><td><button class="brsnr-btn brsnr-btn-sm" data-view>View</button></td></tr>'
    );
  }

  function pendRow(r) {
    return (
      rowOpen(r) +
      "<td>" +
      R.idCell(r) +
      "</td><td>" +
      U.esc(r.source) +
      "</td><td>" +
      U.val(r.current_hub) +
      '</td><td class="brsnr-num">' +
      U.num(r.aging) +
      '</td><td class="brsnr-num">' +
      R.days(r) +
      '</td><td class="brsnr-num">' +
      U.money(r.total_price) +
      "</td><td>" +
      R.risk(r.risk_level) +
      "</td></tr>"
    );
  }

  function newRow(r) {
    return (
      rowOpen(r) +
      "<td>" +
      R.idCell(r) +
      "</td><td>" +
      U.esc(r.source) +
      "</td><td>" +
      U.val(r.current_hub) +
      "</td><td>" +
      U.val(r.partner) +
      '</td><td class="brsnr-num">' +
      U.money(r.total_price) +
      "</td><td>" +
      U.date(r.brsnr_date) +
      "</td></tr>"
    );
  }

  function clearedRow(r) {
    return (
      rowOpen(r) +
      "<td>" +
      R.idCell(r) +
      "</td><td>" +
      U.esc(r.source) +
      "</td><td>" +
      U.val(r.current_hub) +
      '</td><td class="brsnr-num">' +
      U.num(r.aging) +
      '</td><td class="brsnr-num">' +
      U.money(r.total_price) +
      "</td><td>" +
      U.date(r.brsnr_date) +
      "</td></tr>"
    );
  }

  function lossRow(r) {
    return (
      rowOpen(r) +
      "<td>" +
      R.idCell(r) +
      "</td><td>" +
      U.esc(r.source) +
      "</td><td>" +
      U.val(r.current_hub) +
      "</td><td>" +
      U.month(r.brsnr_month) +
      "</td><td>" +
      U.month(r.due_month) +
      "</td><td>" +
      U.date(r.loss_date) +
      '</td><td class="brsnr-num">' +
      U.money(r.total_price) +
      '</td><td class="brsnr-num"><strong>' +
      U.money(r.loss_amount) +
      "</strong></td></tr>"
    );
  }

  var RISK_HEAD =
    "<tr>" +
    "<th>Shipment</th>" +
    "<th>Source</th>" +
    "<th>Current Hub</th>" +
    '<th class="brsnr-num">Aging</th>' +
    '<th class="brsnr-num">Days left</th>' +
    '<th class="brsnr-num">Total price</th>' +
    "<th>Risk</th>" +
    "</tr>";

  var LOSS_HEAD =
    "<tr>" +
    "<th>Shipment ID</th>" +
    "<th>Source</th>" +
    "<th>Current Hub</th>" +
    "<th>BRSNR Month</th>" +
    "<th>Due Month</th>" +
    "<th>Loss Date</th>" +
    '<th class="brsnr-num">Total Price</th>' +
    '<th class="brsnr-num">Loss Amount</th>' +
    "</tr>";

  /* =============================== PAGER =============================== */

  var PAGERS = {
    "brsnr-pager-shipments":
      function () {
        return S.ship;
      },

    "brsnr-pager-risk":
      function () {
        return S.risk;
      }
  };

  function renderPager(
    id,
    st
  ) {
    var el = byId(id);

    var pages = Math.max(
      1,
      Math.ceil(
        st.total / st.size
      )
    );

    var from = st.total
      ? (st.page - 1) *
          st.size +
        1
      : 0;

    var to = Math.min(
      st.total,
      st.page * st.size
    );

    el.innerHTML =
      "<span>" +
      U.num(from) +
      "–" +
      U.num(to) +
      " of " +
      U.num(st.total) +
      '</span><div class="brsnr-pager-ctl">' +
      '<select class="brsnr-select-sm" data-ps aria-label="Rows per page">' +
      [25, 50, 100]
        .map(function (n) {
          return (
            '<option value="' +
            n +
            '"' +
            (n === st.size
              ? " selected"
              : "") +
            ">" +
            n +
            " / page</option>"
          );
        })
        .join("") +
      "</select>" +
      '<button class="brsnr-btn brsnr-btn-sm" data-pg="prev"' +
      (st.page <= 1
        ? " disabled"
        : "") +
      ">Prev</button>" +
      "<span>Page " +
      st.page +
      " / " +
      pages +
      "</span>" +
      '<button class="brsnr-btn brsnr-btn-sm" data-pg="next"' +
      (st.page >= pages
        ? " disabled"
        : "") +
      ">Next</button>" +
      "</div>";
  }

  /* =============================== DRAWER =============================== */

  var Drawer = {
    lastFocus: null,

    open: function (r) {
      if (!r) return;

      Drawer.lastFocus =
        document.activeElement;

      var title =
        byId("brsnr-drawer-title");

      var sub =
        byId("brsnr-drawer-sub");

      var body =
        byId("brsnr-drawer-body");

      if (title) {
        title.textContent =
          r.shipment_id || "Shipment";
      }

      if (sub) {
        sub.textContent =
          (r.source || "—") +
          " • " +
          (r.current_hub ||
            "Current hub unavailable");
      }

      var rows = [
        ["Shipment ID", r.shipment_id],
        ["Casper ID", r.casper_id],
        ["Profile ID", r.profile_id],
        ["Partner", r.partner],
        ["Source", r.source],
        ["State", r.state],
        ["Source Type", r.source_type],
        ["Source Zone", r.source_zone],
        ["Current Hub", r.current_hub],
        ["Delivery Hub", r.delivery_hub],
        ["Delivery Pincode", r.delivery_pincode],
        ["BRSNR Status", R.status(r.brsnr_status), true],
        ["Clearance", R.clearance(r.clearance_status), true],
        ["Risk", R.risk(r.risk_level), true],
        ["BRSNR Date", U.date(r.brsnr_date)],
        ["BRSNR Month", U.month(r.brsnr_month)],
        ["Due Month", U.month(r.due_month)],
        ["Loss Date", U.date(r.loss_date)],
        ["Total Price", U.moneyOrDash(r.total_price)],
        ["Loss Amount", U.moneyOrDash(r.loss_amount)],
        ["Days Remaining", R.days(r)],
        ["Aging", U.num(r.aging)],
        ["Aging Bucket", r.aging_bucket],
        ["First Receive Hub", r.first_receive_hub],
        ["First Receive Time", U.dtime(r.first_receive_time)],
        ["Last Receive Time", U.dtime(r.last_receive_time)],
        ["Previous Hub", r.previous_hub],
        ["Previous Status", r.previous_status],
        ["Final Remarks", r.final_remarks]
      ];

      body.innerHTML =
        R.dl(rows);

      var drawer =
        byId("brsnr-drawer");
      var root =
        byId("brsnr-drawer-root");

      if (!drawer || !root) return;

      root.hidden = false;

      requestAnimationFrame(
        function () {
          root.classList.add(
            "brsnr-open"
          );

          var close =
            drawer.querySelector(
              "[data-drawer-close]"
            );

          if (close) {
            close.focus();
          }
        }
      );
    },

    close: function () {
      var root = byId("brsnr-drawer-root");

      if (!root || root.hidden) return;

      root.classList.remove(
        "brsnr-open"
      );

      setTimeout(function () {
        root.hidden = true;
      }, 220);

      if (
        Drawer.lastFocus &&
        typeof Drawer.lastFocus.focus ===
          "function"
      ) {
        Drawer.lastFocus.focus();
      }
    }
  };

  /* =============================== DATA HELPERS =============================== */

  function dashArgs(extra) {
    var a = {
      hub: S.f.hub,
      from_date: S.f.from,
      to_date: S.f.to,
      status: S.f.status,
      clearance_status:
        S.f.clearance,
      search: S.f.search,
      limit: MAX_PAGE_LIMIT,
      start: 0
    };

    Object.keys(extra || {}).forEach(
      function (k) {
        a[k] = extra[k];
      }
    );

    return a;
  }

  function listArgs(
    page,
    size,
    sort,
    dir,
    extra
  ) {
    return dashArgs(
      Object.assign(
        {
          limit: size,
          start:
            (page - 1) * size,
          sort_by:
            SORTABLE[sort]
              ? sort
              : "aging",
          sort_order:
            dir === "asc"
              ? "asc"
              : "desc"
        },
        extra || {}
      )
    );
  }

  function dataKey() {
    return JSON.stringify({
      hub: S.f.hub,
      from: S.f.from,
      to: S.f.to,
      status: S.f.status,
      clearance: S.f.clearance,
      search: S.f.search
    });
  }

  function fetchBase(args) {
    var key =
      dataKey() +
      "|" +
      JSON.stringify(
        args || {}
      );

    if (
      S.loaded[key] &&
      S.loaded[key].data
    ) {
      return Promise.resolve(
        S.loaded[key].data
      );
    }

    return Api.post(
      "get_dashboard_data",
      args || dashArgs({})
    ).then(function (d) {
      S.loaded[key] = {
        data: d
      };

      if (
        d &&
        d.latest_brsnr_date
      ) {
        S.latest =
          d.latest_brsnr_date;
      }

      if (
        d &&
        d.records
      ) {
        reg(d.records);
      }

      updateNavCounts(d);
      return d;
    });
  }

  function needBase() {
    var key =
      "BASE|" +
      dataKey();

    if (
      S.loaded[key] &&
      S.loaded[key].data
    ) {
      return Promise.resolve(
        S.loaded[key].data
      );
    }

    return Api.post(
      "get_dashboard_data",
      dashArgs({
        limit: 1,
        start: 0
      })
    ).then(function (d) {
      S.loaded[key] = {
        data: d
      };

      S.data = d;

      updateNavCounts(d);

      if (
        d &&
        d.latest_brsnr_date
      ) {
        S.latest =
          d.latest_brsnr_date;
      }

      if (
        d &&
        d.records
      ) {
        reg(d.records);
      }

      return d;
    });
  }

  function more(
    total,
    shown,
    cols
  ) {
    if (
      Number(total || 0) <=
      Number(shown || 0)
    ) {
      return "";
    }

    return (
      '<tr><td colspan="' +
      cols +
      '"><div class="brsnr-table-more">' +
      "Showing " +
      U.num(shown) +
      " of " +
      U.num(total) +
      " authorized records." +
      "</div></td></tr>"
    );
  }

  function riskSplit(d) {
    if (!d) {
      return {
        critical: null,
        watch: null
      };
    }

    var risk =
      d.risk || [];

    var critical =
      risk.filter(function (r) {
        return (
          r.risk_level ===
          "Critical"
        );
      }).length;

    var dueSoon =
      Number(
        d.summary &&
          d.summary.due_soon
      ) || 0;

    /*
      Backend risk endpoint is intentionally capped at 100.
      If it returns fewer than 100 rows, Critical can be counted exactly.
      Watch is derived only when the complete Due Soon population is
      available on the current bounded query.
    */

    var watch = null;

    if (dueSoon <= MAX_PAGE_LIMIT) {
      var allDue =
        (d.records || []).filter(
          function (r) {
            return (
              r.clearance_status ===
              "Due Soon"
            );
          }
        );

      if (
        allDue.length ===
        dueSoon
      ) {
        watch =
          allDue.filter(
            function (r) {
              return (
                r.risk_level ===
                "Watch"
              );
            }
          ).length;
      }
    }

    return {
      critical:
        risk.length >=
          RISK_LIMIT &&
        dueSoon >=
          RISK_LIMIT
          ? null
          : critical,

      watch: watch
    };
  }

  function hubClear(h) {
    var den =
      Number(h.cleared_count || 0) +
      Number(h.pending_count || 0);

    if (!den) return 0;

    return (
      Number(h.cleared_count || 0) /
      den *
      100
    );
  }

  function hubMetricBars(
    id,
    hubs,
    key,
    tone,
    emptyMsg
  ) {
    var labels =
      (hubs || []).map(
        function (h) {
          return h.hub;
        }
      );

    var data =
      (hubs || []).map(
        function (h) {
          return Number(
            h[key] || 0
          );
        }
      );

    Charts.draw(
      id,
      "bar",
      labels,
      [
        {
          label: key,
          data: data,
          brsnrTone: tone
        }
      ],
      {
        emptyMsg:
          emptyMsg ||
          "No data."
      }
    );
  }

  function drawRisk(
    id,
    d
  ) {
    var r =
      d.risk || [];

    var counts = {
      Critical: 0,
      Loss: 0
    };

    r.forEach(function (x) {
      if (
        counts[
          x.risk_level
        ] !== undefined
      ) {
        counts[
          x.risk_level
        ]++;
      }
    });

    Charts.draw(
      id,
      "doughnut",
      ["Critical", "Loss"],
      [
        {
          label: "Risk",
          data: [
            counts.Critical,
            counts.Loss
          ],
          brsnrTones: [
            "crit",
            "loss"
          ]
        }
      ],
      {
        emptyMsg:
          "No Critical or Loss records returned by the risk endpoint."
      }
    );
  }

  /* =============================== OVERVIEW =============================== */

  var L = {};

  L.overview = function () {
    var ok =
      guard("overview");

    R.skelKpis(
      byId(
        "brsnr-kpis"
      ),
      7
    );

    return needBase()
      .then(function (d) {
        if (
          !d ||
          !ok()
        ) {
          return;
        }

        var s =
          d.summary || {};

        S.data = d;

        byId(
          "brsnr-kpis"
        ).innerHTML = [
          R.kpi({
            label: "Total",
            value: U.num(
              s.total
            ),
            sub:
              "Current snapshot"
          }),

          R.kpi({
            label: "New",
            value: U.num(
              s["new"]
            ),
            tone: "info",
            sub:
              "New events in range"
          }),

          R.kpi({
            label: "Pending",
            value: U.num(
              s.pending
            ),
            tone: "warn",
            sub:
              "Open in snapshot"
          }),

          R.kpi({
            label: "Cleared",
            value: U.num(
              s.cleared
            ),
            tone: "ok",
            sub:
              "Cleared events in range"
          }),

          R.kpi({
            label: "Due Soon",
            value: U.num(
              s.due_soon
            ),
            tone: "warn",
            sub:
              "Current snapshot"
          }),

          R.kpi({
            label:
              "Past Due — Loss",
            value: U.num(
              s.loss
            ),
            tone: "loss",
            sub:
              U.money(
                s.loss_value
              )
          }),

          R.kpi({
            label:
              "Clearance %",
            value:
              U.dec(
                s.clearance_percentage,
                1
              ) + "%",
            tone: "ok",
            sub:
              "Cleared ÷ (Cleared + open)"
          })
        ].join("");

        renderOverviewCharts(d);
        renderOverviewPanels(d);
        setLive("live");
      })
      .catch(function (e) {
        if (ok()) {
          fail(e);
        }
      });
  };

  function renderOverviewCharts(
    d
  ) {
    var tr =
      d.trend || [];

    var lab =
      tr.map(function (r) {
        return U.shortDate(
          r.brsnr_date
        );
      });

    Charts.draw(
      "brsnr-ch-trend",
      "bar",
      lab,
      [
        {
          label: "New",
          data: tr.map(
            function (r) {
              return r.new_count;
            }
          ),
          brsnrTone: "info"
        },

        {
          label: "Pending",
          data: tr.map(
            function (r) {
              return r.pending_count;
            }
          ),
          brsnrTone: "warn"
        },

        {
          label: "Cleared",
          data: tr.map(
            function (r) {
              return r.cleared_count;
            }
          ),
          brsnrTone: "ok"
        }
      ],
      {
        emptyMsg:
          "No BRSNR snapshots in this range."
      }
    );

    drawRisk("brsnr-ch-risk", d);
  }

  /* ---- Analytics ---- */

  var ANALYTICS_IDS = [
    "flow",
    "risk",
    "trend",
    "clear",
    "aging",
    "agingtrend",
    "src",
    "cur",
    "value",
    "loss",
    "hubs"
  ];

  L.analytics = function () {
    var ok =
      guard("analytics");

    ANALYTICS_IDS.forEach(
      function (i) {
        Charts.loading(
          "brsnr-ch-an-" + i
        );
      }
    );

    return needBase()
      .then(function (d) {
        if (!d || !ok()) return;

        var tr =
          d.trend || [];

        var lab =
          tr.map(
            function (r) {
              return U.shortDate(
                r.brsnr_date
              );
            }
          );

        var hubs =
          d.hub_performance ||
          [];

        var col =
          function (k) {
            return tr.map(
              function (r) {
                return r[k];
              }
            );
          };

        Charts.draw(
          "brsnr-ch-an-flow",
          "bar",
          lab,
          [
            {
              label: "New",
              data: col(
                "new_count"
              ),
              brsnrTone: "info"
            },

            {
              label: "Pending",
              data: col(
                "pending_count"
              ),
              brsnrTone: "warn"
            },

            {
              label: "Cleared",
              data: col(
                "cleared_count"
              ),
              brsnrTone: "ok"
            }
          ],
          {
            emptyMsg:
              "No BRSNR snapshots in this range."
          }
        );

        drawRisk(
          "brsnr-ch-an-risk",
          d
        );

        Charts.draw(
          "brsnr-ch-an-trend",
          "line",
          lab,
          [
            {
              label:
                "Open BRSNR (New + Pending)",
              brsnrTone:
                "accent",
              fill: true,

              data: tr.map(
                function (r) {
                  return (
                    r.new_count +
                    r.pending_count
                  );
                }
              )
            }
          ],
          {
            emptyMsg:
              "No BRSNR snapshots in this range."
          }
        );

        Charts.draw(
          "brsnr-ch-an-clear",
          "line",
          lab,
          [
            {
              label:
                "Clearance %",
              brsnrTone:
                "ok",
              fill: true,

              data: tr.map(
                function (r) {
                  var den =
                    r.new_count +
                    r.pending_count +
                    r.cleared_count;

                  return den
                    ? +(
                        r.cleared_count /
                        den *
                        100
                      ).toFixed(1)
                    : 0;
                }
              )
            }
          ],
          {
            fmt:
              function (v) {
                return (
                  v + "%"
                );
              },

            max: 100,

            emptyMsg:
              "No clearances in this range."
          }
        );

        Charts.na(
          "brsnr-ch-an-aging",
          "Aging distribution is not provided by the dashboard API."
        );

        Charts.na(
          "brsnr-ch-an-agingtrend",
          "Aging trend is not provided by the dashboard API."
        );

        hubMetricBars(
          "brsnr-ch-an-src",
          hubs,
          "total",
          "info",
          "No source data."
        );

        Charts.na(
          "brsnr-ch-an-cur",
          "Current-hub distribution is not provided by the dashboard API."
        );

        Charts.na(
          "brsnr-ch-an-value",
          "Value buckets are not provided by the dashboard API."
        );

        Charts.draw(
          "brsnr-ch-an-loss",
          "line",
          lab,
          [
            {
              label:
                "Cumulative loss value",

              brsnrTone:
                "loss",

              fill: true,

              data: col(
                "loss_value"
              )
            }
          ],
          {
            fmt:
              function (v) {
                return U.moneyShort(
                  v
                );
              },

            emptyMsg:
              "No loss recorded.",

            tooltip: {
              label:
                function (c) {
                  var x =
                    tr[
                      c.dataIndex
                    ] || {};

                  return (
                    " " +
                    U.money(
                      x.loss_value
                    ) +
                    " · " +
                    U.num(
                      x.loss_count
                    ) +
                    " shipments"
                  );
                }
            }
          }
        );

        Charts.draw(
          "brsnr-ch-an-hubs",
          "bar",
          hubs.map(
            function (h) {
              return h.hub;
            }
          ),
          [
            {
              label:
                "Pending",

              data: hubs.map(
                function (h) {
                  return h.pending_count;
                }
              ),

              brsnrTone:
                "warn"
            },

            {
              label:
                "Cleared",

              data: hubs.map(
                function (h) {
                  return h.cleared_count;
                }
              ),

              brsnrTone:
                "ok"
            }
          ],
          {
            emptyMsg:
              "No hub data."
          }
        );

        S.loaded.analytics =
          true;
      })
      .catch(function (e) {
        if (ok()) {
          fail(e);
        }
      });
  };

  /* ---- Hub performance ---- */

  L.hubperf = function () {
    var ok =
      guard("hubperf");

    var tb =
      byId(
        "brsnr-tbody-hubs"
      );

    R.skelRows(
      tb,
      12,
      6
    );

    return needBase()
      .then(function (d) {
        if (!d || !ok())
          return;

        var hubs =
          (
            d.hub_performance ||
            []
          )
            .map(function (h) {
              return {
                h: h,
                p: hubClear(h)
              };
            })
            .sort(
              function (a, b) {
                return (
                  b.p - a.p ||
                  b.h.cleared_count -
                    a.h.cleared_count
                );
              }
            );

        tb.innerHTML =
          hubs.length
            ? hubs
                .map(
                  function (x, i) {
                    var h =
                      x.h;

                    return (
                      "<tr>" +
                      "<td><strong>" +
                      (i + 1) +
                      "</strong></td>" +

                      "<td><strong>" +
                      U.esc(
                        h.hub
                      ) +
                      "</strong>" +

                      (
                        h.hub_name &&
                        h.hub_name !==
                          h.hub
                          ? '<div class="brsnr-muted">' +
                            U.esc(
                              h.hub_name
                            ) +
                            "</div>"
                          : ""
                      ) +

                      "</td>" +

                      '<td class="brsnr-num">' +
                      U.num(
                        h.total
                      ) +
                      "</td>" +

                      '<td class="brsnr-num">' +
                      U.num(
                        h.new_count
                      ) +
                      "</td>" +

                      '<td class="brsnr-num">' +
                      U.num(
                        h.pending_count
                      ) +
                      "</td>" +

                      '<td class="brsnr-num">' +
                      U.num(
                        h.cleared_count
                      ) +
                      "</td>" +

                      '<td class="brsnr-num brsnr-muted">—</td>' +

                      '<td class="brsnr-num">' +
                      (
                        h.loss_count
                          ? R.badge(
                              U.num(
                                h.loss_count
                              ),
                              "loss"
                            )
                          : "0"
                      ) +
                      "</td>" +

                      '<td style="min-width:150px">' +
                      '<div style="display:flex;align-items:center;gap:8px">' +
                      R.bar(
                        x.p,
                        R.clearBarTone(
                          x.p
                        )
                      ) +
                      "<strong>" +
                      U.dec(
                        x.p,
                        1
                      ) +
                      "%</strong></div>" +
                      "</td>" +

                      '<td class="brsnr-num">' +
                      U.dec(
                        h.avg_aging,
                        1
                      ) +
                      "</td>" +

                      '<td class="brsnr-num">' +
                      U.money(
                        h.loss_value
                      ) +
                      "</td>" +

                      '<td class="brsnr-muted">—</td>' +

                      "</tr>"
                    );
                  }
                )
                .join("")
            : R.emptyRow(
                12,
                "No hub data",
                "Upload a BRSNR snapshot to see hub performance."
              );

        S.loaded.hubperf =
          true;
      })
      .catch(function (e) {
        if (ok()) {
          fail(
            e,
            tb,
            12
          );
        }
      });
  };

  /* ---- History ---- */

  function fillDates(
    sel,
    dates,
    value
  ) {
    sel.innerHTML =
      dates
        .map(function (d) {
          return (
            '<option value="' +
            U.esc(d) +
            '">' +
            U.date(d) +
            "</option>"
          );
        })
        .join("");

    if (value) {
      sel.value = value;
    }
  }

  function snapshotArgs(d) {
    return {
      hub: S.f.hub,
      from_date: d,
      to_date: d,
      limit: 1
    };
  }

  function loadSnapshotKpis() {
    var d =
      byId(
        "brsnr-hist-date"
      ).value;

    var box =
      byId(
        "brsnr-hist-kpis"
      );

    if (!d) {
      return Promise.resolve();
    }

    var ok =
      guard("histkpi");

    R.skelKpis(
      box,
      7
    );

    return Api.post(
      "get_dashboard_data",
      snapshotArgs(d)
    ).then(function (o) {
      if (!ok()) return;

      var s =
        o.summary;

      box.innerHTML = [
        R.kpi({
          label: "Total",
          value: U.num(
            s.total
          ),
          sub:
            "BRSNR rows"
        }),

        R.kpi({
          label: "New",
          value: U.num(
            s["new"]
          )
        }),

        R.kpi({
          label: "Pending",
          value: U.num(
            s.pending
          ),
          tone: "warn"
        }),

        R.kpi({
          label: "Cleared",
          value: U.num(
            s.cleared
          ),
          tone: "ok"
        }),

        R.kpi({
          label:
            "Clearance %",
          value:
            U.dec(
              s.clearance_percentage,
              1
            ) + "%",
          tone: "ok"
        }),

        R.kpi({
          label: "Due Soon",
          value: U.num(
            s.due_soon
          ),
          tone: "warn"
        }),

        R.kpi({
          label:
            "Loss (cumulative)",
          value: U.num(
            s.loss
          ),
          tone: "loss",
          sub:
            U.money(
              s.loss_value
            )
        })
      ].join("");
    });
  }

  L.history = function () {
    var ok =
      guard("history");

    var anchor =
      U.parse(S.latest) ||
      new Date();

    var from =
      new Date(anchor);

    from.setDate(
      anchor.getDate() - 365
    );

    return Api.post(
      "get_dashboard_data",
      {
        hub: S.f.hub,
        from_date:
          U.iso(from),
        to_date:
          S.f.to,
        limit: 1
      }
    )
      .then(function (r) {
        if (!ok()) return;

        if (
          r.latest_brsnr_date
        ) {
          S.latest =
            r.latest_brsnr_date;
        }

        var seen = {};

        var dates =
          (r.trend || [])
            .map(
              function (t) {
                return String(
                  t.brsnr_date
                );
              }
            )
            .filter(
              function (x) {
                if (seen[x])
                  return false;

                seen[x] = 1;
                return true;
              }
            )
            .reverse();

        S.dates = dates;

        if (!dates.length) {
          byId(
            "brsnr-hist-kpis"
          ).innerHTML =
            '<div class="brsnr-empty"><strong>No snapshots yet</strong></div>';

          setLive("live");

          return;
        }

        fillDates(
          byId(
            "brsnr-hist-date"
          ),
          dates,
          dates[0]
        );

        fillDates(
          byId(
            "brsnr-cmp-a"
          ),
          dates,
          dates[1] ||
            dates[0]
        );

        fillDates(
          byId(
            "brsnr-cmp-b"
          ),
          dates,
          dates[0]
        );

        S.loaded.history =
          true;

        return loadSnapshotKpis()
          .then(
            function () {
              if (ok()) {
                setLive("live");
              }
            }
          );
      })
      .catch(function (e) {
        if (ok()) {
          fail(e);
        }
      });
  };

  function delta(
    a,
    b,
    dec
  ) {
    var x =
      Number(b || 0) -
      Number(a || 0);

    return (
      (x > 0 ? "+" : "") +
      (
        dec
          ? U.dec(x, dec)
          : U.num(x)
      )
    );
  }

  function runCompare() {
    var a =
      byId(
        "brsnr-cmp-a"
      ).value;

    var b =
      byId(
        "brsnr-cmp-b"
      ).value;

    var box =
      byId(
        "brsnr-cmp-result"
      );

    if (!a || !b) {
      toast(
        "Pick two dates first.",
        "error"
      );

      return;
    }

    var ok =
      guard("compare");

    box.innerHTML =
      '<span class="brsnr-skel" style="height:90px"></span>';

    Promise.all([
      Api.post(
        "get_dashboard_data",
        snapshotArgs(a)
      ),

      Api.post(
        "get_dashboard_data",
        snapshotArgs(b)
      )
    ])
      .then(function (res) {
        if (!ok()) return;

        var A =
          res[0].summary;

        var B =
          res[1].summary;

        var h =
          '<div class="brsnr-kpi-grid">' +
          [
            R.kpi({
              label: "Total",
              value: U.num(
                B.total
              ),
              sub:
                "A " +
                U.num(A.total) +
                " → B " +
                U.num(B.total) +
                " (" +
                delta(
                  A.total,
                  B.total
                ) +
                ")"
            }),

            R.kpi({
              label: "New",
              value: U.num(
                B["new"]
              ),
              sub:
                "A " +
                U.num(
                  A["new"]
                ) +
                " → B " +
                U.num(
                  B["new"]
                ) +
                " (" +
                delta(
                  A["new"],
                  B["new"]
                ) +
                ")"
            }),

            R.kpi({
              label: "Pending",
              value: U.num(
                B.pending
              ),
              tone: "warn",
              sub:
                "A " +
                U.num(
                  A.pending
                ) +
                " → B " +
                U.num(
                  B.pending
                ) +
                " (" +
                delta(
                  A.pending,
                  B.pending
                ) +
                ")"
            }),

            R.kpi({
              label: "Cleared",
              value: U.num(
                B.cleared
              ),
              tone: "ok",
              sub:
                "A " +
                U.num(
                  A.cleared
                ) +
                " → B " +
                U.num(
                  B.cleared
                ) +
                " (" +
                delta(
                  A.cleared,
                  B.cleared
                ) +
                ")"
            }),

            R.kpi({
              label: "Due Soon",
              value: U.num(
                B.due_soon
              ),
              tone: "warn",
              sub:
                "A " +
                U.num(
                  A.due_soon
                ) +
                " → B " +
                U.num(
                  B.due_soon
                ) +
                " (" +
                delta(
                  A.due_soon,
                  B.due_soon
                ) +
                ")"
            }),

            R.kpi({
              label: "Loss",
              value: U.num(
                B.loss
              ),
              tone: "loss",
              sub:
                "A " +
                U.num(
                  A.loss
                ) +
                " → B " +
                U.num(
                  B.loss
                ) +
                " (" +
                delta(
                  A.loss,
                  B.loss
                ) +
                ")"
            }),

            R.kpi({
              label:
                "Loss Value",
              value:
                U.moneyShort(
                  B.loss_value
                ),
              tone: "loss",
              title:
                U.money(
                  B.loss_value
                ),
              sub:
                "A " +
                U.moneyShort(
                  A.loss_value
                ) +
                " → B " +
                U.moneyShort(
                  B.loss_value
                )
            }),

            R.kpi({
              label:
                "Clearance %",
              value:
                U.dec(
                  B.clearance_percentage,
                  1
                ) +
                "%",
              tone: "ok",
              sub:
                "A " +
                U.dec(
                  A.clearance_percentage,
                  1
                ) +
                "% → B " +
                U.dec(
                  B.clearance_percentage,
                  1
                ) +
                "% (" +
                delta(
                  A.clearance_percentage,
                  B.clearance_percentage,
                  1
                ) +
                ")"
            })
          ].join("") +
          "</div>";

        var map = {};

        (
          res[0]
            .hub_performance ||
          []
        ).forEach(
          function (x) {
            map[x.hub] = {
              a: x
            };
          }
        );

        (
          res[1]
            .hub_performance ||
          []
        ).forEach(
          function (x) {
            (
              map[x.hub] =
                map[x.hub] ||
                {}
            ).b = x;
          }
        );

        var blank = {
          pending_count: 0,
          cleared_count: 0,
          loss_count: 0
        };

        var rows =
          Object.keys(map)
            .sort()
            .map(
              function (hub) {
                var x =
                  map[hub].a ||
                  blank;

                var y =
                  map[hub].b ||
                  blank;

                return (
                  "<tr>" +
                  "<td><strong>" +
                  U.esc(hub) +
                  "</strong></td>" +

                  '<td class="brsnr-num">' +
                  U.num(
                    x.pending_count
                  ) +
                  " → " +
                  U.num(
                    y.pending_count
                  ) +
                  " (" +
                  delta(
                    x.pending_count,
                    y.pending_count
                  ) +
                  ")</td>" +

                  '<td class="brsnr-num">' +
                  U.num(
                    x.cleared_count
                  ) +
                  " → " +
                  U.num(
                    y.cleared_count
                  ) +
                  " (" +
                  delta(
                    x.cleared_count,
                    y.cleared_count
                  ) +
                  ")</td>" +

                  '<td class="brsnr-num">' +
                  U.num(
                    x.loss_count
                  ) +
                  " → " +
                  U.num(
                    y.loss_count
                  ) +
                  " (" +
                  delta(
                    x.loss_count,
                    y.loss_count
                  ) +
                  ")</td>" +

                  "</tr>"
                );
              }
            )
            .join("");

        h +=
          '<div class="brsnr-table-wrap" style="max-height:340px">' +
          '<table class="brsnr-table">' +
          "<thead>" +
          "<tr>" +
          "<th>Source hub</th>" +
          '<th class="brsnr-num">Pending A → B</th>' +
          '<th class="brsnr-num">Cleared A → B</th>' +
          '<th class="brsnr-num">Loss A → B</th>' +
          "</tr>" +
          "</thead>" +
          "<tbody>" +
          (
            rows ||
            R.emptyRow(
              4,
              "No hub data",
              ""
            )
          ) +
          "</tbody>" +
          "</table>" +
          "</div>";

        box.innerHTML = h;
      })
      .catch(function (e) {
        if (!ok()) return;

        box.innerHTML = "";

        fail(e);
      });
  }

  /* ---- Reports ---- */

  var COLS_RECORDS = [
    ["Shipment ID", "shipment_id"],
    ["Casper ID", "casper_id"],
    ["Partner", "partner"],
    ["Source", "source"],
    ["Current Hub", "current_hub"],
    ["BRSNR Status", "brsnr_status"],
    ["Clearance Status", "clearance_status"],
    ["Risk", "risk_level"],
    ["Aging", "aging"],
    ["Days Remaining", "days_remaining"],
    ["Total Price", "total_price"],
    ["Loss Amount", "loss_amount"],
    ["BRSNR Date", "brsnr_date"],
    ["BRSNR Month", "brsnr_month"],
    ["Due Month", "due_month"]
  ];

  var COLS_LOSS = [
    ["Shipment ID", "shipment_id"],
    ["Source", "source"],
    ["Current Hub", "current_hub"],
    ["BRSNR Month", "brsnr_month"],
    ["Due Month", "due_month"],
    ["Loss Date", "loss_date"],
    ["Total Price", "total_price"],
    ["Loss Amount", "loss_amount"]
  ];

  var COLS_RISK = [
    ["Shipment ID", "shipment_id"],
    ["Source", "source"],
    ["Current Hub", "current_hub"],
    ["Aging", "aging"],
    ["Days Remaining", "days_remaining"],
    ["Total Price", "total_price"],
    ["Risk", "risk_level"],
    ["Clearance Status", "clearance_status"],
    ["Loss Amount", "loss_amount"]
  ];

  var COLS_HUB = [
    ["Hub", "hub"],
    ["Hub Name", "hub_name"],
    ["Total", "total"],
    ["New", "new_count"],
    ["Pending", "pending_count"],
    ["Cleared", "cleared_count"],
    ["Loss", "loss_count"],
    ["Avg Aging", "avg_aging"],
    ["Loss Value", "loss_value"]
  ];

  L.reports = function () {
    var hub =
      S.f.hub ||
      (
        S.session &&
        S.session.role ===
          "Director"
          ? "all hubs"
          : "your assigned hubs"
      );

    byId(
      "brsnr-report-note"
    ).textContent =
      "Reports use the current hub filter (" +
      hub +
      ") and the snapshot date resolved from the To date (latest if empty). " +
      "Rows are fetched from the dashboard API in authorized pages. " +
      "Historical-comparison export is not available; use the comparison in History.";

    return needBase()
      .catch(function (e) {
        fail(e);
      });
  };

  function fetchAllRecords(
    extra
  ) {
    var all = [];
    var start = 0;
    var last = null;

    function next() {
      return Api.post(
        "get_dashboard_data",
        dashArgs(
          Object.assign(
            {
              limit:
                MAX_PAGE_LIMIT,
              start: start
            },
            extra || {}
          )
        )
      ).then(
        function (d) {
          last = d;

          all =
            all.concat(
              d.records || []
            );

          start +=
            MAX_PAGE_LIMIT;

          if (
            (d.records || [])
              .length <
              MAX_PAGE_LIMIT ||
            all.length >=
              d.records_total ||
            all.length >=
              MAX_REPORT_ROWS
          ) {
            return {
              d: last,
              rows:
                all.slice(
                  0,
                  MAX_REPORT_ROWS
                ),
              truncated:
                all.length >= MAX_REPORT_ROWS &&
                all.length < Number(d.records_total || 0)
            };
          }

          return next();
        }
      );
    }

    return next();
  }

  function buildReport(report) {
    var key = String(report || "").toLowerCase();

    if (
      key.indexOf("compar") >= 0 ||
      key.indexOf("histor") >= 0
    ) {
      return Promise.reject(
        new ApiError(
          "Historical comparison export is not available. Use the comparison in History.",
          0
        )
      );
    }

    if (
      key.indexOf("loss") >= 0 ||
      key.indexOf("hub") >= 0
    ) {
      return Api.post(
        "get_dashboard_data",
        dashArgs({ limit: 1 })
      ).then(function (d) {
        if (key.indexOf("loss") >= 0) {
          return {
            d: d,
            cols: COLS_LOSS,
            rows: d.loss_register || []
          };
        }

        if (key.indexOf("hub") >= 0) {
          return {
            d: d,
            cols: COLS_HUB,
            rows: d.hub_performance || []
          };
        }

      });
    }

    if (key.indexOf("risk") >= 0) {
      return fetchAllRecords({
        sort_by: "aging",
        sort_order: "desc"
      }).then(function (r) {
        r.cols = COLS_RISK;
        r.rows = r.rows.filter(function (row) {
          return (
            row.risk_level === "Watch" ||
            row.risk_level === "Critical" ||
            row.risk_level === "Loss"
          );
        });
        return r;
      });
    }

    var extra = {};

    if (key.indexOf("pending") >= 0) {
      extra.status = "Pending";
    } else if (key.indexOf("new") >= 0) {
      extra.status = "New";
    } else if (key.indexOf("clear") >= 0) {
      extra.status = "Cleared";
    } else {
      extra.status = S.f.status;
      extra.clearance_status = S.f.clearance;
    }

    extra.sort_by = "aging";
    extra.sort_order = "desc";

    return fetchAllRecords(extra).then(function (r) {
      r.cols = COLS_RECORDS;
      return r;
    });
  }

  function runReport(report, act, btn) {
    btn.disabled = true;

    buildReport(report).then(function (r) {
      var rows = r.rows || [];

      if (!rows.length) {
        toast(
          "This report has no rows for the current selection.",
          "info"
        );
        return;
      }

      if (r.truncated) {
        toast(
          "Report truncated to " +
          U.num(MAX_REPORT_ROWS) +
          " rows.",
          "info"
        );
      }

      var headers = r.cols.map(function (c) {
        return c[0];
      });

      var matrix = rows.map(function (row) {
        return r.cols.map(function (c) {
          return row[c[1]];
        });
      });

      var stamp =
        (r.d && r.d.brsnr_date) ||
        U.iso(new Date());

      if (act === "csv") {
        var csv = [headers]
          .concat(matrix)
          .map(function (x) {
            return x.map(U.csvCell).join(",");
          })
          .join("\r\n");

        var a = document.createElement("a");

        a.href = URL.createObjectURL(
          new Blob(
            ["\ufeff" + csv],
            { type: "text/csv;charset=utf-8" }
          )
        );

        a.download =
          "brsnr-" +
          report +
          "-" +
          stamp +
          ".csv";

        document.body.appendChild(a);
        a.click();
        a.remove();

        setTimeout(function () {
          URL.revokeObjectURL(a.href);
        }, 2000);

        toast(
          "CSV downloaded (" +
          U.num(rows.length) +
          " rows).",
          "ok"
        );
      } else {
        var area = document.createElement("div");
        area.className = "brsnr-print-area";

        area.innerHTML =
          "<h2>BRSNR " +
          U.esc(report) +
          " report · " +
          U.esc(stamp) +
          "</h2><table><thead><tr>" +
          headers.map(function (h) {
            return "<th>" + U.esc(h) + "</th>";
          }).join("") +
          "</tr></thead><tbody>" +
          matrix.map(function (x) {
            return "<tr>" +
              x.map(function (c) {
                return "<td>" + U.val(c) + "</td>";
              }).join("") +
              "</tr>";
          }).join("") +
          "</tbody></table>";

        document.body.appendChild(area);
        document.body.classList.add("brsnr-printing");

        var done = function () {
          document.body.classList.remove("brsnr-printing");
          area.remove();
          window.removeEventListener(
            "afterprint",
            done
          );
        };

        window.addEventListener(
          "afterprint",
          done
        );

        window.print();
      }
    }).catch(function (e) {
      if (e.status !== 401) {
        toast(e.message, "error");
      }
    }).then(function () {
      btn.disabled = false;
    });
  }

  /* ---- Profile / session ---- */
  L.profile = function () {
    return needBase().catch(function (e) {
      fail(e);
    });
  };

  function renderSession() {
    var s = S.session,
      dir = s.role === "Director";

    var roleLabel = dir
      ? "Director"
      : "Hub Incharge";

    var scope = dir
      ? "GLOBAL ACCESS"
      : "ASSIGNED HUB ACCESS";

    byId("brsnr-avatar").textContent =
      (s.name || "?").charAt(0).toUpperCase();

    byId("brsnr-user-name").textContent = s.name;
    byId("brsnr-user-role").textContent = roleLabel;

    byId("brsnr-pop-name").textContent = s.name;
    byId("brsnr-pop-email").textContent = s.email;
    byId("brsnr-pop-role").textContent = roleLabel;

    byId("brsnr-access-role").textContent = roleLabel;

    byId("brsnr-access-scope").textContent =
      dir
        ? "GLOBAL ACCESS"
        : U.num(s.hubs.length) +
          " assigned hub" +
          (s.hubs.length === 1 ? "" : "s");

    byId("brsnr-pop-hubs").innerHTML =
      dir
        ? R.badge(
            "All hubs (" +
            U.num(s.hubs.length) +
            ")",
            "info"
          )
        : (
            s.hubs.map(function (h) {
              return R.badge(h, "muted");
            }).join("") ||
            '<span class="brsnr-muted">No active assignment</span>'
          );

    renderHubOptions(s.hub_options);
    Filters.write();

    var valid = (s.assignments || []).map(function (a) {
      return "<tr><td><strong>" +
        U.esc(a.hub) +
        "</strong></td><td>" +
        (a.from_date
          ? U.date(a.from_date)
          : "—") +
        "</td><td>" +
        (a.to_date
          ? U.date(a.to_date)
          : "No end date") +
        "</td></tr>";
    }).join("");

    byId("brsnr-profile").innerHTML =
      R.dl([
        ["Name", s.name],
        ["Email", s.email],
        ["Role", roleLabel],
        ["Access level", scope]
      ]) +

      '<h4 style="margin:22px 0 10px" class="brsnr-eyebrow">Assigned hubs</h4>' +

      (
        dir
          ? "<p style=\"margin:0\">Global access to all " +
            U.num(s.hubs.length) +
            " hubs (including inactive), all shipments, history, reports and loss data.</p>"

          : '<div class="brsnr-table-wrap" style="margin:0"><table class="brsnr-table"><thead><tr><th>Hub</th><th>Valid from</th><th>Valid until</th></tr></thead><tbody>' +
            (
              valid ||
              '<tr><td colspan="3" class="brsnr-muted">No currently valid assignment.</td></tr>'
            ) +
            "</tbody></table></div>"
      );
  }

  function renderHubOptions(options) {
    var select = byId("brsnr-hub");
    if (!select) return;

    var current = S.f.hub;
    var allLabel =
      S.session && S.session.role === "Director"
        ? "All hubs"
        : "All assigned hubs";

    select.innerHTML =
      '<option value="">' + U.esc(allLabel) + "</option>" +
      (options || []).map(function (o) {
        var name = o.name || o.hub || "";
        return (
          '<option value="' + U.esc(name) + '">' +
          U.esc(o.hub_name || name) +
          "</option>"
        );
      }).join("");

    select.value = current;
  }

  function updateNavCounts(d) {
    var summary = (d && d.summary) || {};
    var values = {
      pending: summary.pending,
      new: summary["new"],
      cleared: summary.cleared,
      loss: summary.loss
    };

    qsa(".brsnr-nav-count[data-count]").forEach(function (el) {
      var value = values[el.getAttribute("data-count")];
      el.textContent =
        value === null || value === undefined
          ? ""
          : U.num(value);
    });

    var snapshot = byId("brsnr-snapshot-date");
    if (snapshot) {
      snapshot.textContent = U.date(d && d.brsnr_date);
    }
  }

  function renderOverviewPanels(d) {
    var riskRows = d.risk || [];
    reg(riskRows);
    reg(d.loss_register || []);

    var attention = byId("brsnr-ov-attention");
    var urgent = riskRows.slice(0, 5);

    attention.innerHTML = urgent.length
      ? urgent.map(function (r) {
          return (
            '<button type="button" class="brsnr-row-item" data-row="' +
            U.esc(rk(r)) +
            '"><div><strong>' +
            U.esc(r.shipment_id || "Shipment") +
            '</strong><span class="brsnr-muted">' +
            U.esc(r.current_hub || r.source || "—") +
            "</span></div><span>" +
            R.risk(r.risk_level) +
            "</span></button>"
          );
        }).join("")
      : '<div class="brsnr-empty"><strong>No urgent records</strong>There are no Critical or Loss records in the API response.</div>';

    var hubs = (d.hub_performance || []).slice(0, 5);
    byId("brsnr-ov-hubs").innerHTML = hubs.length
      ? hubs.map(function (h) {
          var rate = hubClear(h);
          return (
            '<div class="brsnr-row-item"><div><strong>' +
            U.esc(h.hub_name || h.hub) +
            '</strong><span class="brsnr-muted">' +
            U.num(h.pending_count) +
            " pending · " +
            U.num(h.cleared_count) +
            ' cleared</span></div><span class="brsnr-row-end">' +
            R.bar(rate, R.clearBarTone(rate)) +
            U.dec(rate, 1) +
            "%</span></div>"
          );
        }).join("")
      : '<div class="brsnr-empty"><strong>No hub data</strong>Hub ranking is not provided for this selection.</div>';
  }

  var Filters = {
    read: function () {
      S.f.from = byId("brsnr-from").value;
      S.f.to = byId("brsnr-to").value;
      S.f.hub = byId("brsnr-hub").value;
      S.f.status = byId("brsnr-status").value;
      S.f.clearance = byId("brsnr-clearance").value;
      S.f.search = byId("brsnr-search").value.trim();
    },

    write: function () {
      byId("brsnr-from").value = S.f.from;
      byId("brsnr-to").value = S.f.to;
      byId("brsnr-hub").value = S.f.hub;
      byId("brsnr-status").value = S.f.status;
      byId("brsnr-clearance").value = S.f.clearance;
      byId("brsnr-search").value = S.f.search;

      qsa("#brsnr-quick [data-range]").forEach(function (button) {
        button.classList.toggle(
          "brsnr-active",
          button.getAttribute("data-range") === S.f.range
        );
      });
    },

    apply: function () {
      Filters.read();
      S.ship.page = 1;
      S.risk.page = 1;
      invalidate();
      return App.refresh();
    },

    reset: function () {
      S.f = {
        from: "",
        to: "",
        hub: "",
        status: "",
        clearance: "",
        search: "",
        range: ""
      };
      Filters.write();
      S.ship.page = 1;
      S.risk.page = 1;
      invalidate();
      return App.refresh();
    },

    quick: function (range) {
      var now = new Date();
      var from = new Date(now);
      var to = new Date(now);

      S.f.range = range;

      if (range === "custom") {
        Filters.write();
        return;
      }
      if (range === "yesterday") {
        from.setDate(from.getDate() - 1);
        to = new Date(from);
      } else if (range === "7d") {
        from.setDate(from.getDate() - 6);
      } else if (range === "30d") {
        from.setDate(from.getDate() - 29);
      } else if (range === "month") {
        from = new Date(now.getFullYear(), now.getMonth(), 1);
      } else if (range === "prev-month") {
        from = new Date(now.getFullYear(), now.getMonth() - 1, 1);
        to = new Date(now.getFullYear(), now.getMonth(), 0);
      }

      S.f.from = U.iso(from);
      S.f.to = U.iso(to);
      Filters.write();
      Filters.apply();
    }
  };

  function renderKpisFromSummary(target, summary) {
    var s = summary || {};
    byId(target).innerHTML = [
      R.kpi({ label: "Total", value: U.num(s.total), sub: "Current snapshot" }),
      R.kpi({ label: "New", value: U.num(s["new"]), tone: "info" }),
      R.kpi({ label: "Pending", value: U.num(s.pending), tone: "warn" }),
      R.kpi({ label: "Cleared", value: U.num(s.cleared), tone: "ok" }),
      R.kpi({ label: "Due Soon", value: U.num(s.due_soon), tone: "warn" }),
      R.kpi({
        label: "Past Due — Loss",
        value: U.num(s.loss),
        tone: "loss",
        sub: U.money(s.loss_value)
      }),
      R.kpi({
        label: "Clearance %",
        value: U.dec(s.clearance_percentage, 1) + "%",
        tone: "ok",
        sub: "Cleared ÷ (Cleared + open)"
      })
    ].join("");
  }

  function renderListResult(d, tbody, rowFn, cols, pagerId) {
    var rows = d.records || [];
    reg(rows);
    tbody.innerHTML = rows.length
      ? rows.map(rowFn).join("") +
        more(d.records_total, rows.length, cols)
      : R.emptyRow(cols, "No shipments", "No records match the current selection.");
    if (pagerId) {
      var state = pagerId === "brsnr-pager-risk" ? S.risk : S.ship;
      state.total = Number(d.records_total || 0);
      renderPager(pagerId, state);
    }
  }

  L.shipments = function () {
    var ok = guard("shipments");
    var tbody = byId("brsnr-tbody-shipments");
    R.skelRows(tbody, 12);
    return fetchBase(listArgs(
      S.ship.page,
      S.ship.size,
      S.ship.sort,
      S.ship.dir
    )).then(function (d) {
      if (!d || !ok()) return;
      renderListResult(d, tbody, shipRow, 12, "brsnr-pager-shipments");
      byId("brsnr-ship-count").textContent =
        U.num(d.records_total) + " matching shipments";
      qsa("#brsnr-tbl-shipments th[data-sort]").forEach(function (th) {
        var key = th.getAttribute("data-sort");
        th.setAttribute(
          "data-dir",
          key === S.ship.sort ? S.ship.dir : ""
        );
      });
      S.loaded.shipments = true;
      setLive("live");
    }).catch(function (e) {
      if (ok()) fail(e, tbody, 12);
    });
  };

  L.pending = function () {
    var ok = guard("pending");
    var tbody = byId("brsnr-tbody-pending");
    var parts = S.pendSort.split(":");
    R.skelKpis(byId("brsnr-pending-kpis"), 7);
    R.skelRows(tbody, 7);
    return fetchBase(listArgs(1, 50, parts[0], parts[1], {
      status: "Pending"
    })).then(function (d) {
      if (!d || !ok()) return;
      renderKpisFromSummary("brsnr-pending-kpis", d.summary);
      renderListResult(d, tbody, pendRow, 7);
      Charts.na(
        "brsnr-ch-pending-aging",
        "Aging distribution is not provided as an aggregate by the dashboard API."
      );
      hubMetricBars(
        "brsnr-ch-pending-hubs",
        d.hub_performance,
        "pending_count",
        "warn",
        "No hub data."
      );
      S.loaded.pending = true;
      setLive("live");
    }).catch(function (e) {
      if (ok()) fail(e, tbody, 7);
    });
  };

  L.new = function () {
    var ok = guard("new");
    var tbody = byId("brsnr-tbody-new");
    R.skelKpis(byId("brsnr-new-kpis"), 7);
    R.skelRows(tbody, 6);
    return fetchBase(listArgs(1, 50, "shipment_id", "asc", {
      status: "New"
    })).then(function (d) {
      if (!d || !ok()) return;
      renderKpisFromSummary("brsnr-new-kpis", d.summary);
      renderListResult(d, tbody, newRow, 6);
      hubMetricBars(
        "brsnr-ch-new-src",
        d.hub_performance,
        "new_count",
        "info",
        "No source-hub data."
      );
      Charts.na(
        "brsnr-ch-new-cur",
        "Current-hub distribution is not provided as an aggregate by the dashboard API."
      );
      S.loaded.new = true;
      setLive("live");
    }).catch(function (e) {
      if (ok()) fail(e, tbody, 6);
    });
  };

  L.cleared = function () {
    var ok = guard("cleared");
    var tbody = byId("brsnr-tbody-cleared");
    R.skelKpis(byId("brsnr-cleared-kpis"), 7);
    R.skelRows(tbody, 6);
    return fetchBase(listArgs(1, 50, "shipment_id", "asc", {
      status: "Cleared"
    })).then(function (d) {
      if (!d || !ok()) return;
      renderKpisFromSummary("brsnr-cleared-kpis", d.summary);
      renderListResult(d, tbody, clearedRow, 6);
      var trend = d.trend || [];
      Charts.draw(
        "brsnr-ch-cleared",
        "line",
        trend.map(function (r) { return U.shortDate(r.brsnr_date); }),
        [{
          label: "Cleared",
          data: trend.map(function (r) { return r.cleared_count; }),
          brsnrTone: "ok",
          fill: true
        }],
        { emptyMsg: "No cleared events in this date range." }
      );
      S.loaded.cleared = true;
      setLive("live");
    }).catch(function (e) {
      if (ok()) fail(e, tbody, 6);
    });
  };

  function riskRow(r) {
    return (
      rowOpen(r) +
      "<td>" + R.idCell(r) + "</td>" +
      "<td>" + U.val(r.source) + "</td>" +
      "<td>" + U.val(r.current_hub) + "</td>" +
      '<td class="brsnr-num">' +
      (r.aging === null || r.aging === undefined ? "—" : U.num(r.aging)) +
      "</td>" +
      '<td class="brsnr-num">' + R.days(r) + "</td>" +
      '<td class="brsnr-num">' + U.money(r.total_price) + "</td>" +
      "<td>" + R.risk(r.risk_level || "Loss") + "</td></tr>"
    );
  }

  L.risk = function () {
    var ok = guard("risk");
    var tbody = byId("brsnr-tbody-risk");
    var bucket = S.risk.bucket || "Critical";
    R.skelRows(tbody, 7);
    byId("brsnr-thead-risk").innerHTML = RISK_HEAD;

    return Api.post(
      "get_dashboard_data",
      dashArgs({ limit: MAX_PAGE_LIMIT, start: 0 })
    ).then(function (d) {
      if (!d || !ok()) return;
      reg(d.records);
      reg(d.loss_register);

      var allRows = bucket === "Loss"
        ? (d.loss_register || []).map(function (r) {
            return Object.assign({}, r, { risk_level: "Loss" });
          })
        : (d.records || []).filter(function (r) {
            return r.risk_level === bucket;
          });

      var complete = (d.records || []).length >= Number(d.records_total || 0);
      S.risk.total = allRows.length;
      S.risk.page = Math.min(
        S.risk.page,
        Math.max(1, Math.ceil(S.risk.total / S.risk.size))
      );

      var summary = d.summary || {};
      var counts = {
        Critical: complete
          ? (d.records || []).filter(function (r) { return r.risk_level === "Critical"; }).length
          : "—",
        Watch: complete
          ? (d.records || []).filter(function (r) { return r.risk_level === "Watch"; }).length
          : "—",
        Loss: U.num(summary.loss)
      };
      byId("brsnr-risk-tiles").innerHTML = [
        ["Critical", "crit"],
        ["Watch", "warn"],
        ["Loss", "loss"]
      ].map(function (x) {
        return (
          '<button type="button" class="brsnr-risk-tile' +
          (bucket === x[0] ? " brsnr-active" : "") +
          '" data-tone="' + x[1] +
          '" data-risk-bucket="' + x[0] + '">' +
          '<span class="brsnr-kpi-label">' + x[0] + "</span>" +
          '<span class="brsnr-kpi-value">' + counts[x[0]] + "</span></button>"
        );
      }).join("");

      byId("brsnr-risk-heading").textContent = bucket;
      var pageRows = allRows.slice(
        (S.risk.page - 1) * S.risk.size,
        S.risk.page * S.risk.size
      );
      tbody.innerHTML = pageRows.length
        ? pageRows.map(riskRow).join("")
        : R.emptyRow(7, "No matching shipments", "No " + bucket + " records are available for this selection.");

      byId("brsnr-risk-count").textContent =
        U.num(allRows.length) +
        (bucket === "Loss" ? " loss-register records" : " matching records") +
        (bucket !== "Loss" && !complete
          ? " in the first " + U.num((d.records || []).length) +
            " of " + U.num(d.records_total) + " records"
          : "");

      renderPager("brsnr-pager-risk", S.risk);
      S.loaded.risk = true;
      setLive("live");
    }).catch(function (e) {
      if (ok()) fail(e, tbody, 7);
    });
  };

  /* =============================== APP CORE =============================== */
  var App = {
    loadSection: function (name) {
      if (name === "overview") {
        return L.overview();
      }

      if (S.loaded[name] && name !== "reports") {
        return Promise.resolve();
      }

      return (
        L[name] ||
        function () {
          return Promise.resolve();
        }
      )();
    },

    refresh: function () {
      invalidate();
      S.data = null;
      S.rec = {};
      S.safeCount = null;

      return App.loadSection(S.section);
    }
  };

  var Nav = {
    go: function (name, bucket, fromHistory) {
      var section = SECTIONS[name] ? name : "overview";

      if (bucket) {
        S.risk.bucket = bucket;
        S.risk.page = 1;
      }

      S.section = section;
      byId("brsnr-title").textContent = SECTIONS[section];
      byId("brsnr-filterbar").hidden = !!NO_FILTERBAR[section];

      qsa(".brsnr-section").forEach(function (el) {
        el.hidden = el.getAttribute("data-section") !== section;
      });
      qsa("#brsnr-nav .brsnr-nav-item").forEach(function (el) {
        el.classList.toggle(
          "brsnr-active",
          el.getAttribute("data-section") === section
        );
        el.setAttribute(
          "aria-current",
          el.getAttribute("data-section") === section ? "page" : "false"
        );
      });

      if (fromHistory) {
        window.history.replaceState(
          null,
          "",
          window.location.pathname +
            window.location.search +
            "#" + section
        );
      } else if (window.location.hash !== "#" + section) {
        window.history.pushState(
          null,
          "",
          window.location.pathname +
            window.location.search +
            "#" + section
        );
      }

      byId("brsnr-app").classList.remove("brsnr-menu-open");
      return App.loadSection(section);
    },

    fromHash: function () {
      var section = window.location.hash.replace(/^#/, "");
      Nav.go(SECTIONS[section] ? section : "overview", null, true);
    }
  };

  /* =============================== EVENTS =============================== */
  function bind() {
    // navigation (delegated)
    byId("brsnr-nav").addEventListener("click", function (e) {
      var b = e.target.closest(".brsnr-nav-item");

      if (b) {
        Nav.go(
          b.getAttribute("data-section")
        );
      }
    });

    window.addEventListener(
      "popstate",
      Nav.fromHash
    );

    byId("brsnr-content").addEventListener(
      "click",
      function (e) {
        var rt = e.target.closest(
          "[data-risk-bucket]"
        );

        if (rt) {
          S.risk.bucket =
            rt.getAttribute(
              "data-risk-bucket"
            );
          S.risk.page = 1;
          L.risk();
          return;
        }

        var g = e.target.closest(
          "[data-goto]"
        );

        if (g) {
          Nav.go(
            g.getAttribute("data-goto"),
            g.getAttribute("data-bucket")
          );
          return;
        }

        var rep = e.target.closest(
          ".brsnr-report [data-act]"
        );

        if (rep) {
          runReport(
            rep.closest(
              ".brsnr-report"
            ).getAttribute(
              "data-report"
            ),
            rep.getAttribute("data-act"),
            rep
          );
          return;
        }

        var row = e.target.closest("[data-row]");

        if (row) {
          Drawer.open(
            S.rec[row.getAttribute("data-row")]
          );
        }
      }
    );

    // shipment table sorting (server-side, backend-supported fields only)
    byId("brsnr-tbl-shipments")
      .tHead
      .addEventListener("click", function (e) {
        var th = e.target.closest(
          "th[data-sort]"
        );

        if (!th) return;

        var k = th.getAttribute(
          "data-sort"
        );

        if (!SORTABLE[k]) {
          toast(
            "Sorting by this column is not supported.",
            "info"
          );
          return;
        }

        S.ship.dir =
          S.ship.sort === k &&
          S.ship.dir === "desc"
            ? "asc"
            : "desc";

        S.ship.sort = k;
        S.ship.page = 1;

        L.shipments();
      });

    // pagination (delegated)
    document.addEventListener(
      "click",
      function (e) {
        var b = e.target.closest(
          "[data-pg]"
        );

        if (!b) return;

        var pager = b.closest(
          ".brsnr-pager"
        );

        if (
          !pager ||
          !PAGERS[pager.id]
        ) {
          return;
        }

        var st = PAGERS[pager.id](),
          pages = Math.max(
            1,
            Math.ceil(
              st.total / st.size
            )
          );

        st.page = Math.min(
          pages,
          Math.max(
            1,
            st.page +
              (
                b.getAttribute("data-pg") === "next"
                  ? 1
                  : -1
              )
          )
        );

        (
          pager.id === "brsnr-pager-risk"
            ? L.risk
            : L.shipments
        )();
      }
    );

    document.addEventListener(
      "change",
      function (e) {
        var s = e.target.closest(
          "[data-ps]"
        );

        if (!s) return;

        var pager = s.closest(
          ".brsnr-pager"
        );

        if (
          !pager ||
          !PAGERS[pager.id]
        ) {
          return;
        }

        var st = PAGERS[pager.id]();

        st.size = Math.min(
          parseInt(s.value, 10) || 50,
          MAX_PAGE_LIMIT
        );

        st.page = 1;

        (
          pager.id === "brsnr-pager-risk"
            ? L.risk
            : L.shipments
        )();
      }
    );

    // filters
    byId("brsnr-apply")
      .addEventListener(
        "click",
        Filters.apply
      );

    byId("brsnr-reset")
      .addEventListener(
        "click",
        Filters.reset
      );

    byId("brsnr-refresh")
      .addEventListener(
        "click",
        function () {
          App.refresh();
        }
      );

    byId("brsnr-refresh-top")
      .addEventListener(
        "click",
        function () {
          App.refresh();
        }
      );

    byId("brsnr-banner-retry")
      .addEventListener(
        "click",
        function () {
          App.refresh();
        }
      );

    byId("brsnr-quick")
      .addEventListener(
        "click",
        function (e) {
          var p = e.target.closest(
            "[data-range]"
          );

          if (p) {
            Filters.quick(
              p.getAttribute(
                "data-range"
              )
            );
          }
        }
      );

    [
      "brsnr-from",
      "brsnr-to"
    ].forEach(function (id) {
      byId(id).addEventListener(
        "change",
        function () {
          S.f.range = "custom";
          Filters.read();
          Filters.write();
        }
      );
    });

    var debounced = U.debounce(
      function () {
        if (
          byId("brsnr-search")
            .value
            .trim()
            .length !== 1
        ) {
          Filters.apply();
        }
      },
      500
    );

    byId("brsnr-search")
      .addEventListener(
        "input",
        debounced
      );

    byId("brsnr-search")
      .addEventListener(
        "keydown",
        function (e) {
          if (e.key === "Enter") {
            Filters.apply();
          }
        }
      );

    byId("brsnr-pending-sort")
      .addEventListener(
        "change",
        function (e) {
          S.pendSort =
            e.target.value;
          L.pending();
        }
      );

    // history
    byId("brsnr-hist-date")
      .addEventListener(
        "change",
        function () {
          loadSnapshotKpis()
            .catch(fail);
        }
      );

    byId("brsnr-cmp-run")
      .addEventListener(
        "click",
        runCompare
      );

    // theme, sidebar, menus
    byId("brsnr-theme")
      .addEventListener(
        "click",
        Theme.toggle
      );

    byId("brsnr-collapse")
      .addEventListener(
        "click",
        function () {
          var c =
            byId("brsnr-app")
              .classList
              .toggle(
                "brsnr-collapsed"
              );

          try {
            localStorage.setItem(
              "brsnr_sidebar",
              c ? "1" : "0"
            );
          } catch (e) {
            /* ignore */
          }
        }
      );

    byId("brsnr-menu")
      .addEventListener(
        "click",
        function () {
          byId("brsnr-app")
            .classList
            .toggle(
              "brsnr-menu-open"
            );
        }
      );

    byId("brsnr-backdrop")
      .addEventListener(
        "click",
        function () {
          byId("brsnr-app")
            .classList
            .remove(
              "brsnr-menu-open"
            );
        }
      );

    byId("brsnr-user-btn")
      .addEventListener(
        "click",
        function (e) {
          e.stopPropagation();

          var pop =
            byId("brsnr-user-pop");

          pop.hidden = !pop.hidden;

          byId("brsnr-user-btn")
            .setAttribute(
              "aria-expanded",
              String(!pop.hidden)
            );
        }
      );

    document.addEventListener(
      "click",
      function (e) {
        if (
          !e.target.closest(
            ".brsnr-user"
          )
        ) {
          byId(
            "brsnr-user-pop"
          ).hidden = true;
        }
      }
    );

    byId("brsnr-logout")
      .addEventListener(
        "click",
        Auth.logout
      );

    // drawer
    byId("brsnr-drawer-close")
      .addEventListener(
        "click",
        Drawer.close
      );

    byId("brsnr-drawer-overlay")
      .addEventListener(
        "click",
        Drawer.close
      );

    document.addEventListener(
      "keydown",
      function (e) {
        if (e.key !== "Escape") return;

        if (
          !byId(
            "brsnr-drawer-root"
          ).hidden
        ) {
          Drawer.close();
        } else {
          byId(
            "brsnr-user-pop"
          ).hidden = true;
        }
      }
    );
  }

  /* =============================== INIT =============================== */
  function init() {
    try {
      if (
        localStorage.getItem(
          "brsnr_sidebar"
        ) === "1"
      ) {
        byId("brsnr-app")
          .classList
          .add(
            "brsnr-collapsed"
          );
      }
    } catch (e) {
      /* ignore */
    }

    byId("brsnr-theme")
      .setAttribute(
        "aria-checked",
        String(
          Theme.get() === "dark"
        )
      );

    bind();

    R.skelKpis(
      byId("brsnr-kpis"),
      9
    );

    ["trend", "risk"].forEach(
      function (i) {
        Charts.loading(
          "brsnr-ch-" + i
        );
      }
    );

    setLive("loading");

    Api.post(
      "get_portal_session"
    ).then(function (s) {
      var opts =
        s.hub_options || [];

      S.session = {
        role: s.role,
        name:
          s.incharge_name ||
          s.email,
        email: s.email,
        assignments:
          s.assignments || [],
        hub_options: opts,
        hubs: opts.map(
          function (o) {
            return o.name;
          }
        )
      };

      renderSession();
      Nav.fromHash();
    }).catch(function (e) {
      fail(e);
    });
  }

  init();
});