document.addEventListener('DOMContentLoaded', function () {
  const toggleButtons = document.querySelectorAll('.toggle-password');
  toggleButtons.forEach(function (button) {
    button.addEventListener('click', function () {
      const input = document.getElementById(button.dataset.target);
      if (!input) return;
      const type = input.getAttribute('type') === 'password' ? 'text' : 'password';
      input.setAttribute('type', type);
      button.innerHTML = type === 'password'
        ? '<i class="fa-solid fa-eye"></i>'
        : '<i class="fa-solid fa-eye-slash"></i>';
    });
  });

  const flashMessages = document.querySelectorAll('.flash');
  flashMessages.forEach(function (flash) {
    setTimeout(function () {
      flash.style.opacity = '0';
      flash.style.transform = 'translateY(-5px)';
      flash.style.transition = 'opacity 0.3s ease, transform 0.3s ease';
    }, 4000);
  });
});
