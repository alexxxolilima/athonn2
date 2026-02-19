document.addEventListener("DOMContentLoaded", () => {
  const dateEl = document.getElementById("current-date");
  if (dateEl) {
    const now = new Date();
    const fmt = new Intl.DateTimeFormat("pt-BR", {
      weekday: "long",
      day: "2-digit",
      month: "long",
      year: "numeric"
    });
    dateEl.textContent = `Central de atendimento técnico • ${fmt.format(now)}`;
  }

  document.querySelectorAll(".module-card").forEach((card, idx) => {
    card.style.opacity = "0";
    card.style.transform = "translateY(10px)";
    setTimeout(() => {
      card.style.transition = "opacity 260ms ease, transform 260ms ease";
      card.style.opacity = "1";
      card.style.transform = "translateY(0)";
    }, 80 * (idx + 1));
  });
});
