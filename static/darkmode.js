if (localStorage.getItem('dark-mode') === 'true') {
  document.documentElement.classList.add('dark');
  document.getElementById('dark-toggle').textContent = '☀️ Light';
}


const darkToggle = document.getElementById('dark-toggle');
darkToggle.addEventListener('click', () => {
  document.documentElement.classList.toggle('dark');
  const isDark = document.documentElement.classList.contains('dark');
  localStorage.setItem('dark-mode', isDark);
  darkToggle.textContent = isDark ? '☀️ Light' : '🌙 Dark';
});
