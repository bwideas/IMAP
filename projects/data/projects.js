/*
  PROJECT STORIES — the only file you need to edit to add or update news.

  How to add a project:
    1. Copy one block { ... } below and paste it at the top of the list.
    2. Give it a unique "id" (lowercase, dashes, no spaces) — this becomes the share link:
         story.html?id=your-id
    3. Fill in the fields. Leave a field as "" or [] if you don't have it yet.
    4. Set published: true ONLY after the customer has approved external publication.
    5. Put photos in projects/images/ and reference them as "images/your-photo.jpg".

  Fields marked sample: true show a "Sample" badge — delete these blocks before sharing the link.
*/
window.SITE = {
  name: "CBS Project Stories",
  tagline: "Commercial Building Services — landmark projects across Saudi Arabia & MENA, 2025/2026",
  contactEmail: "your.team@example.com",     // replace with the team mailbox
  contactPhone: "",                          // optional, e.g. "+966 ..."
  contactCta: "Planning a similar project? Talk to our CBS team."
};

window.PROJECTS = [
  {
    id: "sample-data-center-cooling",
    sample: true,
    published: true,
    featured: true,
    title: "Data centre chilled-water system upgrade",
    customer: "[Customer name — with approval]",
    date: "2026-03",
    country: "Saudi Arabia",
    city: "Riyadh",
    industry: "Data Center",
    application: "HVAC",
    image: "",
    summary: "Replace with a 2–3 sentence summary: what was built, for whom, and the headline result.",
    challenge: "Describe the customer's challenge: uptime requirements, energy costs, space limits, tight schedule…",
    solution: "Describe what the team delivered: pump selection, controls, commissioning, service agreement…",
    metrics: [
      { value: "XX%", label: "Energy saving" },
      { value: "XX", label: "Pumps installed" },
      { value: "XX MW", label: "Cooling capacity" }
    ],
    results: [
      "Result 1 — measurable outcome",
      "Result 2 — reliability / uptime improvement",
      "Result 3 — delivered on schedule"
    ],
    products: ["Product / solution 1", "Product / solution 2"],
    quote: { text: "Add an approved customer quote here.", author: "Name", role: "Title, Company" },
    gallery: []
  },
  {
    id: "sample-mixed-use-pressure-boosting",
    sample: true,
    published: true,
    featured: false,
    title: "Mixed-use tower pressure boosting",
    customer: "[Customer name — with approval]",
    date: "2025-11",
    country: "Saudi Arabia",
    city: "Jeddah",
    industry: "Mixed-Use Buildings",
    application: "Pressure boosting",
    image: "",
    summary: "Replace with a short summary of the project and its headline result.",
    challenge: "Describe the challenge.",
    solution: "Describe the solution.",
    metrics: [
      { value: "XX", label: "Floors served" },
      { value: "XX%", label: "Energy saving" }
    ],
    results: ["Result 1", "Result 2"],
    products: ["Product / solution 1"],
    quote: null,
    gallery: []
  },
  {
    id: "sample-hospital-hvac",
    sample: true,
    published: true,
    featured: false,
    title: "Hospital HVAC pumping optimisation",
    customer: "[Customer name — with approval]",
    date: "2026-06",
    country: "Egypt",
    city: "Cairo",
    industry: "Health care",
    application: "Heating & Cooling",
    image: "",
    summary: "Replace with a short summary of the project and its headline result.",
    challenge: "Describe the challenge.",
    solution: "Describe the solution.",
    metrics: [{ value: "XX%", label: "Energy saving" }],
    results: ["Result 1"],
    products: ["Product / solution 1"],
    quote: null,
    gallery: []
  },
  {
    id: "sample-district-cooling",
    sample: true,
    published: true,
    featured: false,
    title: "District cooling plant distribution pumps",
    customer: "[Customer name — with approval]",
    date: "2026-09",
    country: "Saudi Arabia",
    city: "Dammam",
    industry: "District cooling",
    application: "HVAC",
    image: "",
    summary: "Replace with a short summary of the project and its headline result.",
    challenge: "Describe the challenge.",
    solution: "Describe the solution.",
    metrics: [{ value: "XX TR", label: "Plant capacity" }],
    results: ["Result 1"],
    products: ["Product / solution 1"],
    quote: null,
    gallery: []
  },
  {
    id: "example-draft-hidden",
    sample: true,
    published: false,   // waiting for customer approval — NOT shown on the site
    featured: false,
    title: "Draft story awaiting customer approval",
    customer: "",
    date: "2026-10",
    country: "Saudi Arabia",
    city: "",
    industry: "Hospitality",
    application: "Wastewater",
    image: "",
    summary: "",
    challenge: "",
    solution: "",
    metrics: [],
    results: [],
    products: [],
    quote: null,
    gallery: []
  }
];
