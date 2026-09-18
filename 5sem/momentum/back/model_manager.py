import json

from django.core.paginator import Paginator
from django.db.models import F, Sum
from back.forms import LoginForm, RegisterForm, AskForm, AnswerForm, EditProfileForm
from back.models import Post, Answer, Tag, Profile, PostLike, AnswerLike


def get_popular_tags():
    # print(Tag.objects.get_queryset().order_by('used_times')[:10])
    return Tag.objects.get_queryset().order_by('-used_times')[:10]


def get_popular_users():
    # print(Profile.objects.get_queryset().order_by('answers_count')[0].user.username)
    return Profile.objects.get_queryset().order_by('-answers_count')[:5]


def get_answer_form(request):
    if request.method == "POST":
        return AnswerForm(request.POST)
    return AnswerForm()


def get_login_form(request):
    if request.method == "POST":
        return LoginForm(data=request.POST)
    return LoginForm()


def get_signup_form(request):
    if request.method == "POST":
        return RegisterForm(data=request.POST, files=request.FILES)
    return RegisterForm()


def index_news(request):
    user_profile = request.user.profile if request.user.is_authenticated else None
    return Post.objects.news(user_profile=user_profile)


def hot_news(request):
    user_profile = request.user.profile if request.user.is_authenticated else None
    return Post.objects.hot(user_profile=user_profile)


def tag_news(tag):
    return Post.objects.filter(tags__tag_name=tag)



def check_page(request):
    try:
        page = int(request.GET['page'])
        if page < 1:
            return 1
    except:
        return 1

    return page


def pagination(request, type_req, count=4, tag_name=None):
    page_num = check_page(request)
    items = ''
    user_profile = request.user.profile if request.user.is_authenticated else None

    if type_req == 'index':
        items = index_news(request)
    elif type_req == 'tag':
        items = tag_news(tag_name).visible_to_user(user_profile)
    elif type_req == 'hot':
        items = hot_news(request)

    paginator = Paginator(items, count)
    items = paginator.get_page(page_num)

    return {'questions': items, 'paginator': paginator}


def this_question(request, question_id, count=4):
    page_num = check_page(request)
    try:
        question = Post.objects.get(pk=question_id)
    except:
        return {'page': -1}

    tags = question.tags.all()
    answers = Answer.objects.filter(question=question_id).order_by('-likes_count')

    paginator = Paginator(answers, count)
    answers = paginator.get_page(page_num)

    return {
        'question': question, 'tags': tags, 'answers': answers,
        'paginator': paginator
    }


def answer(request, question_id):
    if request.method == "POST":
        form = AnswerForm(request.POST)
        if form.is_valid():
            a = form.save()
            question = Post.objects.get(pk=question_id)
            a.question = question
            a.author = request.user.profile
            Profile.objects.filter(id=request.user.profile.id).update(answers_count=F('answers_count') + 1)
            Post.objects.filter(id=question_id).update(answers_count=F('answers_count') + 1)
            a.save()


def like(request):
    body = json.loads(request.body)
    req_type = body['type']
    flag = body['flag']
    user = request.user.profile

    # print(user)
    __id = body['id']
    if req_type == 'question':
        q = Post.objects.get_queryset().get(pk=__id)
        PostLike.objects.add_vote(user_id=user, question_id=q, vote=flag)
        q.likes_count = PostLike.objects.filter(question_id=q.id).aggregate(Sum('vote'))['vote__sum']
        q.save()
        return PostLike.objects.filter(question_id=__id).aggregate(Sum('vote'))['vote__sum']
    else:
        a = Answer.objects.get_queryset().get(pk=__id)
        AnswerLike.objects.add_vote(user_id=user, answer_id=a, vote=flag)
        a.likes_count = AnswerLike.objects.filter(answer_id=a.id).aggregate(Sum('vote'))['vote__sum']
        a.save()
        print(AnswerLike.objects.filter(answer_id=a.id).aggregate(Sum('vote'))['vote__sum'])
        return AnswerLike.objects.filter(answer_id=a.id).aggregate(Sum('vote'))['vote__sum']


def correct(request):
    body = json.loads(request.body)
    flag = body['flag']
    ans_id = body['id']

    a = Answer.objects.get(pk=ans_id)
    a.correct = flag
    a.save()
    return flag


def get_edit_form(request):
    if request.method == 'POST':
        return EditProfileForm(data=request.POST, files=request.FILES)
    return EditProfileForm()
