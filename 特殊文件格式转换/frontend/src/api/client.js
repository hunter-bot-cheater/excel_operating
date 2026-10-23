const baseURL = import.meta.env.VITE_API_BASE_URL || "/api/v1";

export async function convertSpecialFormat(formData) {
  const response = await fetch(`${baseURL}/special-format-conversions/convert`, { method: "POST", body: formData });
  return response.json();
}

export async function getTask(taskId) {
  const response = await fetch(`${baseURL}/special-format-conversions/tasks/${taskId}`);
  return response.json();
}
