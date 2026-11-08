const navButtons = document.querySelectorAll(".nav-item");
const views = {
  "whole-document": document.querySelector("#view-whole-document"),
  "part-document": document.querySelector("#view-part-document"),
  "kb-correction": document.querySelector("#view-kb-correction"),
  "docx-convert": document.querySelector("#view-docx-convert"),
};

navButtons.forEach((button) => {
  button.addEventListener("click", () => {
    const target = button.dataset.view;
    navButtons.forEach((b) => b.classList.toggle("active", b === button));
    Object.entries(views).forEach(([key, el]) => {
      el.hidden = key !== target;
    });
  });
});
