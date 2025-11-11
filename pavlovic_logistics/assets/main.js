const navToggle = document.querySelector('.menu-toggle');
const navList = document.querySelector('#nav-list');
const yearSpan = document.querySelector('#year');

if (yearSpan) {
    yearSpan.textContent = new Date().getFullYear();
}

if (navToggle && navList) {
    navToggle.addEventListener('click', () => {
        const isOpen = navList.classList.toggle('open');
        navToggle.setAttribute('aria-expanded', String(isOpen));
    });
}

const navLinks = document.querySelectorAll('#nav-list a');
navLinks.forEach((link) =>
    link.addEventListener('click', () => {
        navList.classList.remove('open');
        navToggle.setAttribute('aria-expanded', 'false');
    })
);
