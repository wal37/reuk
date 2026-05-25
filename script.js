const navToggle = document.querySelector(".nav-toggle");
const mainNav = document.querySelector(".main-nav");

if (navToggle && mainNav) {
  navToggle.addEventListener("click", () => {
    const open = mainNav.classList.toggle("is-open");
    navToggle.setAttribute("aria-expanded", String(open));
  });
}

for (const form of document.querySelectorAll("form")) {
  form.addEventListener("submit", (event) => {
    event.preventDefault();
    const button = form.querySelector("button[type='submit']");
    if (!button) return;
    const original = button.textContent;
    button.textContent = "Submitted";
    button.disabled = true;
    setTimeout(() => {
      button.textContent = original;
      button.disabled = false;
      form.reset();
    }, 1800);
  });
}

const articleCarousel = document.querySelector("[data-article-carousel]");

if (articleCarousel) {
  const track = articleCarousel.querySelector(".article-track");
  const pages = Array.from(articleCarousel.querySelectorAll(".article-page"));
  const prev = articleCarousel.querySelector(".article-arrow-left");
  const next = articleCarousel.querySelector(".article-arrow-right");
  let index = 0;

  const render = () => {
    track.style.transform = `translateX(-${index * 100}%)`;
  };

  prev?.addEventListener("click", () => {
    index = (index - 1 + pages.length) % pages.length;
    render();
  });

  next?.addEventListener("click", () => {
    index = (index + 1) % pages.length;
    render();
  });
}

const institutionToggle = document.querySelector("[data-institution-toggle]");
const institutionWall = document.querySelector("[data-institution-wall]");
const expandableInstitutions = Array.from(document.querySelectorAll(".institution-badge.is-expandable"));

if (institutionToggle && institutionWall && expandableInstitutions.length) {
  institutionToggle.addEventListener("click", () => {
    const expanded = institutionToggle.getAttribute("aria-expanded") === "true";
    institutionWall.setAttribute("data-expanded", String(!expanded));
    institutionToggle.setAttribute("aria-expanded", String(!expanded));
    institutionToggle.textContent = expanded ? "Show more" : "Show less";
  });
}
