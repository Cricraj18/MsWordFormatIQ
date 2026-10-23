import { getTitleText } from "./version";

// Placeholder page. The real interface is built in Task 7.
const app = document.getElementById("app");
if (app) {
  app.textContent = getTitleText();
}
