// Please see documentation at https://learn.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.
document.addEventListener('DOMContentLoaded', function () {
    // 사이드바 링크들
    var links = document.querySelectorAll('#sidebar-accordion .nav-link');

    links.forEach(function (a) {
        // 기존 클릭 포커스 제거(있으면 유지)
        a.addEventListener('click', function () {
            a.blur();
        });

        // 마우스 들어오면 배경 ON
        a.addEventListener('mouseenter', function () {
            a.classList.add('is-hover');
        });

        // 마우스 나가면 배경 OFF (무조건 원복)
        a.addEventListener('mouseleave', function () {
            a.classList.remove('is-hover');
        });
    });
});