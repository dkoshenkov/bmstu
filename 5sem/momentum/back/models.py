from datetime import timedelta
from django.db.models import Q
from django.db import models
from django.contrib.auth.models import User
from django.db.models import Sum
from django.utils import timezone


# =================Tag============================== #

class TagQuerySet(models.QuerySet):
    def is_tag(self, tag):
        return self.filter(tag_name=tag)

    def increment_used_times(self, tag_ids):
        self.filter(id__in=tag_ids).update(used_times=models.F('used_times') + 1)


class TagManager(models.Manager):
    def get_queryset(self):
        return TagQuerySet(self.model, using=self._db)

    def is_tag(self, tag):
        return self.get_queryset().is_tag(tag)


class Tag(models.Model):
    tag_name = models.CharField(max_length=32, unique=True)
    used_times = models.PositiveIntegerField(default=0)

    objects = TagManager()

    def __str__(self):
        return self.tag_name


# =================Profile============================== #

class ProfileQuerySet(models.QuerySet):
    def is_profile(self, profile):
        return self.filter(user__username=profile)

    def increment_answers_count(self, profile_id):
        self.filter(id=profile_id).update(answers_count=(models.F('answers_count') + 1))


class ProfileManager(models.Manager):
    def get_queryset(self):
        return ProfileQuerySet(self.model, using=self._db)

    def is_profile(self, profile):
        return self.get_queryset().is_profile(profile)


class Profile(models.Model):
    updated_time = models.DateTimeField(auto_now=True)
    avatar_path = models.ImageField(upload_to='images', default='default3.jpg')
    answers_count = models.PositiveIntegerField(default=0)

    user = models.OneToOneField(User, on_delete=models.CASCADE)
    objects = ProfileManager()

    def recalculate_answers_count(self):
        self.answers_count = Answer.objects.filter(author=self).count()
        self.save()

    def __str__(self):
        return self.user.username

    def __int__(self):
        return self.pk


# ============Question=================================== #

class PostQueryset(models.QuerySet):
    def order_rating(self):
        return self.order_by('-created_time', '-likes_count')

    def order_time(self):
        return self.order_by('-created_time')

    def tags(self, tag_name):
        return self.filter(tags__tag_name=tag_name).order_by('created_time')

    def get_tags(self):
        return [tag.tag_name for tag in self.tags.all()]

    def increment_answers_count(self, question_id):
        # Увеличиваем счетчик ответов для вопроса
        self.filter(id=question_id).update(answers_count=models.F('answers_count') + 1)

    def visible_to_user(self, user_profile):
        # Returns posts based on the user's subscription status
        if user_profile:
            return self.filter(
                Q(access_mode='public') |
                Q(access_mode='subscribers', author__followers__follower=user_profile)
            ).distinct()
        else:
            # For unauthenticated users, only return public posts
            return self.filter(access_mode='public')


class PostManager(models.Manager):
    def get_queryset(self):
        return PostQueryset(self.model, using=self._db)

    def hot(self, user_profile=None):
        one_week_ago = timezone.now() - timedelta(days=7)
        return self.get_queryset().visible_to_user(user_profile).order_by('-likes_count').filter(created_time__gte=one_week_ago)

    def news(self, user_profile=None):
        return self.get_queryset().visible_to_user(user_profile).order_by('-created_time')

    def get_tags(self):
        return self.get_queryset().get_tags()


class Post(models.Model):
    content = models.TextField()
    author = models.ForeignKey(Profile, on_delete=models.SET_NULL, null=True)
    created_time = models.DateTimeField(default=timezone.now)
    likes_count = models.IntegerField(default=0)
    answers_count = models.IntegerField(default=0)

    tags = models.ManyToManyField(Tag, related_name='questions')

    objects = PostManager()

    ACCESS_CHOICES = [
        ('public', 'Общий доступ'),
        ('subscribers', 'Только для подписчиков'),
    ]
    access_mode = models.CharField(max_length=12, choices=ACCESS_CHOICES, default='public')

    def __int__(self):
        return self.pk

    def save(self, *args, **kwargs):
        super().save(*args, **kwargs)

        if self.pk:
            Tag.objects.get_queryset().increment_used_times(self.tags.values_list('id', flat=True))

    def recalculate_likes(self):
        # Пересчитываем количество лайков на основе всех записей QuestionLike для данного Question
        self.likes_count = PostLike.objects.filter(question_id=self).aggregate(Sum('vote'))['vote__sum'] or 0
        self.save()

    def recalculate_answers_count(self):
        self.answers_count = self.answer_set.count()
        self.save()


# ===============QuestionLike================================ #

class PostLikeQuerySet(models.QuerySet):
    def add_vote(self, user_id, question_id, vote):
        q = Post.objects.get(pk=question_id)
        try:
            self.create(user_id=user_id, question_id=question_id, vote=vote)
            q.likes_count += vote
            q.save()
            return q.likes_count

        except:
            item = PostLike.objects.get_queryset().get(user_id=user_id, question_id=question_id)
            if item.vote != vote:
                q.likes_count += vote
                item.vote = vote
                item.save()
                q.save()
            else:
                item.vote = 0
                item.save()
                q.likes_count = PostLike.objects.filter(question_id=q.id).aggregate(Sum('vote'))['vote__sum']
                q.save()


class PostLikeManager(models.Manager):
    def get_queryset(self):
        return PostLikeQuerySet(self.model, using=self._db)

    def add_vote(self, user_id, question_id, vote):
        self.get_queryset().add_vote(user_id=user_id, question_id=question_id, vote=vote)


class PostLike(models.Model):
    user_id = models.ForeignKey(Profile, on_delete=models.CASCADE)
    question_id = models.ForeignKey(Post, on_delete=models.CASCADE)
    vote = models.IntegerField()

    objects = PostLikeManager()

    class Meta:
        unique_together = (('question_id', 'user_id'),)


# =============Answer================================== #

class AnswerQuerySet(models.QuerySet):
    def hots(self):
        return self.order_by('-created_time').order_by('-likes_count')


class AnswerManager(models.Manager):
    def get_queryset(self):
        return AnswerQuerySet(self.model, using=self._db)

    def hots(self):
        return self.get_queryset().hots()


class Answer(models.Model):
    objects = AnswerManager()
    answer = models.TextField()
    created_time = models.DateTimeField(auto_now_add=True)
    author = models.ForeignKey(Profile, on_delete=models.SET_NULL, null=True)

    likes_count = models.IntegerField(default=0)
    question = models.ForeignKey(Post, on_delete=models.CASCADE, blank=True, null=True)

    def __str__(self):
        return self.answer

    def __int__(self):
        return self.pk

    def save(self, *args, **kwargs):
        created = self._state.adding

        super().save(*args, **kwargs)

        if created:
            if self.question:
                self.question.answers_count += 1
                self.question.save()
            if self.author:
                self.author.answers_count += 1
                self.author.save()

    def recalculate_likes(self):
        # Пересчитываем количество лайков на основе всех записей AnswerLike для данного Answer
        self.likes_count = AnswerLike.objects.filter(answer=self).aggregate(Sum('vote'))['vote__sum'] or 0
        self.save()


# =============AnsweLiker================================== #

class AnswerLikeQuerySet(models.QuerySet):
    def add_vote(self, user_id, answer_id, vote):
        a = Answer.objects.get(pk=answer_id)
        try:
            item = self.create(user_id=user_id, answer_id=answer_id, vote=vote)
            a.likes_count += vote
            a.save()
            item.save()
        except:
            item = AnswerLike.objects.get_queryset().get(user_id=user_id, answer_id=answer_id)
            if item.vote != vote:
                item.vote = vote
                item.save()

                a.likes_count = AnswerLike.objects.filter(answer_id=a.id).aggregate(Sum('vote'))['vote__sum']
                a.save()
            else:
                item.vote = 0
                item.save()
                a.likes_count = AnswerLike.objects.filter(answer_id=a.id).aggregate(Sum('vote'))['vote__sum']
                a.save()


class AnswerLikeManager(models.Manager):
    def get_queryset(self):
        return AnswerLikeQuerySet(self.model, using=self._db)

    def add_vote(self, user_id, answer_id, vote):
        self.get_queryset().add_vote(user_id=user_id, answer_id=answer_id, vote=vote)


class AnswerLike(models.Model):
    answer = models.ForeignKey(Answer, on_delete=models.CASCADE)
    vote = models.IntegerField()
    user_id = models.ForeignKey(Profile, on_delete=models.CASCADE)

    objects = AnswerLikeManager()

    class Meta:
        unique_together = ('user_id', 'answer')


class PostFile(models.Model):
    post = models.ForeignKey(Post, related_name='files', on_delete=models.CASCADE)
    file = models.FileField(upload_to='uploads/')

    def __str__(self):
        return self.file.name


class Subscription(models.Model):
    follower = models.ForeignKey(Profile, related_name='following', on_delete=models.CASCADE)
    followed = models.ForeignKey(Profile, related_name='followers', on_delete=models.CASCADE)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        unique_together = ('follower', 'followed')