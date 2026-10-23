const baseURL = import.meta.env.VITE_API_BASE_URL || "/api/v1";

export async function screenNonconforming(formData) {
  const response = await fetch(`${baseURL}/nonconforming/screen`, { method: "POST", body: formData });
  return response.json();
}

export async function getTask(taskId) {
  const response = await fetch(`${baseURL}/nonconforming/tasks/${taskId}`);
  return response.json();
}
