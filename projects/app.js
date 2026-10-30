(function(){
  const SITE = window.SITE || {};
  const ALL = (window.PROJECTS || []).filter(p => p.published);
  ALL.sort((a,b) => (b.date || "").localeCompare(a.date || ""));

  const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c =>
    ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
  const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const fmtDate = d => { if(!d) return ""; const [y,m] = d.split("-"); return m ? MONTHS[+m-1] + " " + y : y; };
  const year = p => (p.date || "").slice(0,4);
  const uniq = arr => [...new Set(arr.filter(Boolean))].sort();
  const $ = id => document.getElementById(id);

  function media(p, cls){
    const bg = p.image ? ` style="background-image:url('${esc(p.image)}')"` : "";
    const ph = p.image ? "" : `<div class="ph">${esc(p.industry || "Project")}<br>photo coming soon</div>`;
    const badge = p.sample ? `<span class="badge-sample">Sample</span>` : "";
    return `<div class="media ${cls || ""}"${bg}>${ph}${badge}</div>`;
  }
  function tags(p){
    return `<div class="meta">
      ${p.industry ? `<span class="tag">${esc(p.industry)}</span>` : ""}
      ${p.application ? `<span class="tag">${esc(p.application)}</span>` : ""}
      <span class="tag loc">${esc([p.city, p.country].filter(Boolean).join(", "))}${p.date ? " · " + fmtDate(p.date) : ""}</span>
    </div>`;
  }
  const link = p => "story.html?id=" + encodeURIComponent(p.id);

  function common(){
    document.querySelectorAll("[data-site]").forEach(el => { el.textContent = SITE[el.dataset.site] || ""; });
    document.querySelectorAll("[data-contact]").forEach(el => {
      el.href = "mailto:" + (SITE.contactEmail || "") + "?subject=" + encodeURIComponent(el.dataset.contact || "Project enquiry");
    });
    const y = $("year"); if(y) y.textContent = new Date().getFullYear();
  }

  /* ---------- Home page ---------- */
  function home(){
    // Stats
    $("statProjects").textContent = ALL.length;
    $("statIndustries").textContent = uniq(ALL.map(p => p.industry)).length;
    $("statCities").textContent = uniq(ALL.map(p => p.city)).length;
    $("statCountries").textContent = uniq(ALL.map(p => p.country)).length;

    // Featured
    const feat = ALL.find(p => p.featured);
    if(feat){
      $("featured").innerHTML = `<a class="featured" href="${link(feat)}">
        ${media(feat)}
        <div class="body">${tags(feat)}<h2>${esc(feat.title)}</h2><p>${esc(feat.summary)}</p>
          <p style="margin-top:1rem"><span class="btn">Read the story →</span></p></div></a>`;
    }

    // Filters
    const fill = (id, vals) => { const s = $(id); vals.forEach(v => s.insertAdjacentHTML("beforeend", `<option>${esc(v)}</option>`)); };
    fill("fIndustry", uniq(ALL.map(p => p.industry)));
    fill("fApp", uniq(ALL.map(p => p.application)));
    fill("fCountry", uniq(ALL.map(p => p.country)));
    fill("fYear", uniq(ALL.map(year)).reverse());

    const params = new URLSearchParams(location.search);
    ["industry","app","country","year"].forEach(k => {
      const el = $("f" + k[0].toUpperCase() + k.slice(1));
      if(params.get(k)) el.value = params.get(k);
    });

    function render(){
      const q = $("q").value.trim().toLowerCase();
      const fi = $("fIndustry").value, fa = $("fApp").value, fc = $("fCountry").value, fy = $("fYear").value;
      const list = ALL.filter(p =>
        (!fi || p.industry === fi) && (!fa || p.application === fa) &&
        (!fc || p.country === fc) && (!fy || year(p) === fy) &&
        (!q || [p.title,p.summary,p.city,p.country,p.industry,p.application,p.customer].join(" ").toLowerCase().includes(q)));
      $("count").textContent = list.length + (list.length === 1 ? " project" : " projects");
      $("grid").innerHTML = list.length ? list.map(p => `<a class="card" href="${link(p)}">
          ${media(p)}
          <div class="body">${tags(p)}<h3>${esc(p.title)}</h3><p>${esc(p.summary)}</p><span class="more">Read more →</span></div>
        </a>`).join("") : `<div class="empty">No projects match these filters.</div>`;
    }
    ["q","fIndustry","fApp","fCountry","fYear"].forEach(id => $(id).addEventListener("input", render));
    render();
  }

  /* ---------- Story page ---------- */
  function story(){
    const id = new URLSearchParams(location.search).get("id");
    const p = ALL.find(x => x.id === id);
    if(!p){
      $("story").innerHTML = `<div class="wrap" style="padding:4rem 16px"><div class="empty">
        This story is not available. <a href="index.html">See all projects</a></div></div>`;
      return;
    }
    document.title = p.title + " — " + (SITE.name || "Project Stories");

    const url = location.href;
    const shareText = p.title + " — " + url;
    const metrics = (p.metrics || []).map(m => `<div class="kpi"><b>${esc(m.value)}</b><span>${esc(m.label)}</span></div>`).join("");
    const list = arr => (arr && arr.length) ? `<ul>${arr.map(x => `<li>${esc(x)}</li>`).join("")}</ul>` : "";
    const section = (title, html) => html ? `<section><h2>${title}</h2>${html}</section>` : "";
    const facts = [["Customer",p.customer],["Location",[p.city,p.country].filter(Boolean).join(", ")],
      ["Industry",p.industry],["Application",p.application],["Completed",fmtDate(p.date)]]
      .filter(f => f[1]).map(f => `<dt>${f[0]}</dt><dd>${esc(f[1])}</dd>`).join("");

    $("story").innerHTML = `
      <div class="story-hero"><div class="wrap">
        <a class="back" href="index.html">← All projects</a>
        <h1>${esc(p.title)}</h1>${tags(p)}
        ${p.sample ? `<div class="notice-sample">Sample story — replace with real, customer-approved content before sharing.</div>` : ""}
      </div></div>
      <div class="wrap">
        ${p.image ? media(p, "story-media") : ""}
        ${metrics ? `<div class="kpis">${metrics}</div>` : ""}
        <div class="story-grid">
          <div>
            ${p.summary ? `<section><p class="lead">${esc(p.summary)}</p></section>` : ""}
            ${section("The challenge", p.challenge ? `<p>${esc(p.challenge)}</p>` : "")}
            ${section("Our solution", p.solution ? `<p>${esc(p.solution)}</p>` : "")}
            ${section("Results", list(p.results))}
            ${p.quote && p.quote.text ? `<section><blockquote>“${esc(p.quote.text)}”
              <cite>— ${esc(p.quote.author)}${p.quote.role ? ", " + esc(p.quote.role) : ""}</cite></blockquote></section>` : ""}
            ${(p.gallery || []).length ? section("Gallery", `<div class="gallery">${p.gallery.map(g =>
                `<img src="${esc(g)}" alt="${esc(p.title)}" loading="lazy">`).join("")}</div>`) : ""}
          </div>
          <aside>
            ${facts ? `<div class="box"><h3>Project facts</h3><dl>${facts}</dl></div>` : ""}
            ${(p.products || []).length ? `<div class="box"><h3>Solutions supplied</h3>${list(p.products)}</div>` : ""}
            <div class="box"><h3>Share this project</h3><div class="share">
              <a class="btn" target="_blank" rel="noopener" href="https://wa.me/?text=${encodeURIComponent(shareText)}">WhatsApp</a>
              <a class="btn" target="_blank" rel="noopener" href="https://www.linkedin.com/sharing/share-offsite/?url=${encodeURIComponent(url)}">LinkedIn</a>
              <a class="btn" href="mailto:?subject=${encodeURIComponent(p.title)}&body=${encodeURIComponent(shareText)}">Email</a>
              <button class="btn ghost" id="copyLink">Copy link</button>
            </div></div>
            <div class="box"><h3>Similar project?</h3><p style="margin-bottom:.8rem">${esc(SITE.contactCta || "")}</p>
              <a class="btn" data-contact="Enquiry: ${esc(p.title)}" href="#">Contact us</a></div>
          </aside>
        </div>
      </div>`;

    $("copyLink").addEventListener("click", e => {
      const done = () => { e.target.textContent = "Copied ✓"; setTimeout(() => e.target.textContent = "Copy link", 1500); };
      if(navigator.clipboard) navigator.clipboard.writeText(url).then(done, () => prompt("Copy this link:", url));
      else prompt("Copy this link:", url);
    });
  }

  if(document.body.dataset.page === "story") story(); else home();
  common();
})();
