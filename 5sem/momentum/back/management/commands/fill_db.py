from django.core.management.base import BaseCommand
from back.models import Profile, Post, Answer, Tag, AnswerLike, PostLike, Subscription
from django.contrib.auth.models import User
import random
from faker import Faker
from datetime import timedelta
from django.utils import timezone

fake = Faker('ru_RU')


class Command(BaseCommand):
    help = 'Fill database with test data'

    def add_arguments(self, parser):
        parser.add_argument('ratio', type=int, help='Ratio of data to generate')

    def handle(self, *args, **options):
        ratio = options['ratio']

        # Создание пользователей
        users = []
        for i in range(ratio):
            num = fake.random_number()
            user = User(username=fake.user_name() + f"{num}", email=f"{num}" + fake.email())
            user.set_password('123456')
            users.append(user)
        User.objects.bulk_create(users)

        profiles = [Profile(user=user) for user in User.objects.all()]
        Profile.objects.bulk_create(profiles)

        # Создание тегов
        tags = set()
        existing_tags = set(Tag.objects.values_list('tag_name', flat=True))

        while len(tags) < ratio:
            name = fake.word()
            if name not in existing_tags:
                tags.add(name)

        tag_objects = [Tag(tag_name=name) for name in tags]
        Tag.objects.bulk_create(tag_objects)

        profiles = list(Profile.objects.all())
        posts = []

        # Генерация постов с рандомной датой в пределах последнего месяца
        for i in range(ratio * 10):
            profile = random.choice(profiles)
            random_days_ago = random.randint(0, 30)
            random_hours_ago = random.randint(0, 23)
            created_time = timezone.now() - timedelta(days=random_days_ago, hours=random_hours_ago)
            posts.append(Post(content=fake.paragraph(), author=profile, created_time=created_time))
        Post.objects.bulk_create(posts)

        # Добавление тегов к постам
        posts = list(Post.objects.all())
        tags = list(Tag.objects.all())
        for post in posts:
            random_tags = random.sample(tags, random.randint(1, 5))
            post.tags.set(random_tags)

        # Создание ответов
        answers = []
        for i in range(ratio * 100):
            post = random.choice(posts)
            profile = random.choice(profiles)
            answers.append(Answer(answer=fake.paragraph(), question=post, author=profile))
        Answer.objects.bulk_create(answers)

        # Пересчитываем количество ответов для вопросов
        for post in Post.objects.all():
            post.recalculate_answers_count()

        # Пересчитываем количество ответов для профилей
        for profile in Profile.objects.all():
            profile.recalculate_answers_count()

        # Создание лайков к постам
        post_likes_set = set()
        post_likes = []
        for i in range(ratio * 200):
            post = random.choice(posts)
            profile = random.choice(profiles)
            key = (post.id, profile.id)
            if key not in post_likes_set:
                post_likes_set.add(key)
                vote = random.choice([-1, 1])
                post_likes.append(PostLike(user_id=profile, question_id=post, vote=vote))
        PostLike.objects.bulk_create(post_likes)

        # Создание лайков к ответам
        answer_likes_set = set()
        answer_likes = []
        for i in range(ratio * 200):
            answer = random.choice(answers)
            profile = random.choice(profiles)
            key = (answer.id, profile.id)
            if key not in answer_likes_set:
                answer_likes_set.add(key)
                vote = random.choice([-1, 1])
                answer_likes.append(AnswerLike(user_id=profile, answer_id=answer, vote=vote))
        AnswerLike.objects.bulk_create(answer_likes)

        # Пересчитываем рейтинг для каждого поста
        for post in Post.objects.all():
            post.recalculate_likes()

        # Пересчитываем рейтинг для каждого ответа
        for answer in Answer.objects.all():
            answer.recalculate_likes()
        #
        profiles = Profile.objects.all()  # Получаем все профили
        subscriptions = []

        for follower in profiles:
            for followed in profiles:
                if follower != followed:  # Исключаем подписку на самого себя
                    subscriptions.append(Subscription(follower=follower, followed=followed))

        Subscription.objects.bulk_create(subscriptions)

        self.stdout.write(self.style.SUCCESS(f'Database filled with {ratio}x data'))
