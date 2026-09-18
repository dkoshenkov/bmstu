document.addEventListener('DOMContentLoaded', function () {
    const answer_cards = document.querySelectorAll('.answer_card');
    const question_cards = document.querySelectorAll('.question-card');

    function getCookie(name) {
        let cookieValue = null;
        if (document.cookie && document.cookie !== '') {
            const cookies = document.cookie.split(';');
            for (let i = 0; i < cookies.length; i++) {
                const cookie = cookies[i].trim();
                // Does this cookie string begin with the name we want?
                if (cookie.substring(0, name.length + 1) === (name + '=')) {
                    cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
                    break;
                }
            }
        }
        return cookieValue;
    }

    for (const card of answer_cards) {
        const like_btn = card.querySelector('.btn-group button[type="button"][class*="answer-like-btn"]');
        const dislike_btn = card.querySelector('.btn-group button[type="button"][class*="answer-dislike-btn"]');
        const correct_btn = card.querySelector('.correct-selector');

        if (like_btn) {
            like_btn.addEventListener('click', () => {
                let ans_id = card.getAttribute('id');
                ans_id = Number(ans_id);
                const answer_rating = card.querySelector('.answer-rating');
                const request = new Request('/like/', {
                    method: 'post',
                    headers: {
                        'Content-Type': 'application/json',
                        'X-CSRFToken': getCookie('csrftoken'),
                    },
                    body: JSON.stringify({
                        id: ans_id,
                        type: 'answer',
                        flag: 1,
                    })
                });
                fetch(request)
                    .then((response) => response.json())
                    .then((data) =>
                        answer_rating.innerHTML = data.count);
            });
        }

        if (dislike_btn) {
            dislike_btn.addEventListener('click', () => {
                let ans_id = card.getAttribute('id');
                ans_id = Number(ans_id);
                const answer_rating = card.querySelector('.answer-rating');
                const request = new Request('/like/', {
                    method: 'post',
                    headers: {
                        'Content-Type': 'application/json',
                        'X-CSRFToken': getCookie('csrftoken'),
                    },
                    body: JSON.stringify({
                        id: ans_id,
                        type: 'answer',
                        flag: -1,
                    })
                });
                fetch(request)
                    .then((response) => response.json())
                    .then((data) =>
                        answer_rating.innerHTML = data.count);
            });
        }

        if (correct_btn) {
            correct_btn.addEventListener('click', () => {
                let ans_id = card.getAttribute('id');
                ans_id = Number(ans_id);
                const request = new Request('/correct/', {
                    method: 'post',
                    headers: {
                        'Content-Type': 'application/json',
                        'X-CSRFToken': getCookie('csrftoken'),
                    },
                    body: JSON.stringify({
                        id: ans_id,
                        flag: correct_btn.checked,
                    })
                });
                fetch(request)
                    .then((response) => response.json())
                    .then((data) => {
                        correct_btn.checked = data.flag;
                        if (data.flag === true) {
                            card.classList.add('border-success');
                        } else {
                            card.classList.remove('border-success');
                        }
                    });
            });
        }
    }

    for (const card of question_cards) {
        const like_btn = card.querySelector('.btn-group button[type="button"][class*="question-like-btn"]');
        const dislike_btn = card.querySelector('.btn-group button[type="button"][class*="question-dislike-btn"]');

        if (like_btn) {
            like_btn.addEventListener('click', () => {
                let q_id = card.id;
                q_id = Number(q_id);
                const question_rating = card.querySelector('.question-rating');
                const request = new Request('/like/', {
                    method: 'post',
                    headers: {
                        'Content-Type': 'application/json',
                        'X-CSRFToken': getCookie('csrftoken'),
                    },
                    body: JSON.stringify({
                        id: q_id,
                        type: 'question',
                        flag: 1,
                    })
                });
                fetch(request)
                    .then((response) => response.json())
                    .then((data) =>
                        question_rating.innerHTML = data.count);
            });
        }

        if (dislike_btn) {
            dislike_btn.addEventListener('click', () => {
                let q_id = card.id;
                q_id = Number(q_id);
                const question_rating = card.querySelector('.question-rating');
                const request = new Request('/like/', {
                    method: 'post',
                    headers: {
                        'Content-Type': 'application/json',
                        'X-CSRFToken': getCookie('csrftoken'),
                    },
                    body: JSON.stringify({
                        id: q_id,
                        type: 'question',
                        flag: -1,
                    })
                });
                fetch(request)
                    .then((response) => response.json())
                    .then((data) =>
                        question_rating.innerHTML = data.count);
            });
        }
    }
});