from django import forms
from django.core.validators import RegexValidator, MaxLengthValidator
from django.utils import timezone

from back.models import User, Profile, Post, Tag, Answer, PostFile
from django.core.exceptions import ValidationError

from django.forms import modelformset_factory


class LoginForm(forms.Form):
    username = forms.CharField()
    password = forms.CharField(widget=forms.PasswordInput, min_length=4)

    def clean_username(self):
        try:
            username = self.cleaned_data.get('username')
            user = User.objects.get(username=username)
            return user
        except:
            raise ValidationError('There is no selected user')

    def save(self):
        username = self.cleaned_data.get('username')
        return User.objects.get(username=username)


class RegisterForm(forms.Form):
    email = forms.EmailField(required=True, widget=forms.TextInput(attrs={'class': 'form-control'}))
    username = forms.CharField(
        required=True,
        widget=forms.TextInput(attrs={'class': 'form-control'}),
        validators=[RegexValidator(regex='^[0-9a-zA-Z]*$', message='Username can only contain letters and digits')]
    )
    avatar = forms.ImageField(required=False, widget=forms.FileInput(attrs={'class': 'form-control'}))

    password = forms.CharField(widget=forms.PasswordInput, min_length=4, required=True)
    repeat_password = forms.CharField(widget=forms.PasswordInput, min_length=4, required=True)

    def clean(self):
        super(RegisterForm, self).clean()
        password = self.cleaned_data.get('password')
        repeat_password = self.cleaned_data.get('repeat_password')
        if password != repeat_password:
            raise ValidationError('Passwords do not match')

    def clean_username(self):
        username = self.cleaned_data.get('username')
        if User.objects.filter(username=username).exists():
            raise ValidationError("Username is already taken")
        return username

    def is_password_valid(self):
        super(RegisterForm, self).clean()
        password = self.cleaned_data.get('password')
        repeat_password = self.cleaned_data.get('repeat_password')
        if password != repeat_password:
            return False
        return True

    def is_user_valid(self):
        super(RegisterForm, self).clean()
        email = self.cleaned_data.get('email')
        username = self.cleaned_data.get('username')

        if User.objects.filter(username=username).count() > 0 or User.objects.filter(email=email).count() > 0:
            return False

        return True

    def save(self):
        username = self.cleaned_data.get('username')
        email = self.cleaned_data.get('email')
        password = self.cleaned_data.get('password')
        avatar = self.cleaned_data.get('avatar')
        repeat_password = self.cleaned_data.get('repeat_password')

        if password != repeat_password:
            raise ValidationError('Passwords do not match')

        user = User.objects.create_user(
            username=username,
            email=email,
            password=password)
        if avatar:
            Profile.objects.create(user=user, avatar_path=avatar)
        else:
            Profile.objects.create(user=user)
        return user


class AskForm(forms.ModelForm):
    # Добавляем ограничение длины content до 1000 символов
    content = forms.CharField(
        widget=forms.Textarea,
        validators=[MaxLengthValidator(1000)]
    )

    tags = forms.CharField(
        required=True,
        label='Теги',
        widget=forms.TextInput(attrs={'placeholder': 'Введите теги через пробел'}),
    )

    class Meta:
        model = Post
        fields = ['content', 'tags', 'access_mode']

    def __init__(self, *args, **kwargs):
        super(AskForm, self).__init__(*args, **kwargs)
        if self.instance and self.instance.pk:
            tag_names = ' '.join([tag.tag_name for tag in self.instance.tags.all()])
            self.fields['tags'].initial = tag_names
            # Debug statement
            print("Debug - Initial tags set as plain text:", self.fields['tags'].initial)

    def save(self, commit=True, author=None):
        instance = super().save(commit=False)
        if author:
            instance.author = author
        if commit:
            instance.save()
            instance.tags.clear()
            tags = self.cleaned_data.get('tags').split()
            for t in tags:
                tag_obj, created = Tag.objects.get_or_create(tag_name=t)
                instance.tags.add(tag_obj)
        return instance


PostFileFormSet = modelformset_factory(
    PostFile,
    fields=('file',),
    extra=1,
    max_num=10,
    widgets={'file': forms.FileInput()},
    can_delete=True
)


class AnswerForm(forms.Form):
    text = forms.CharField(
        widget=forms.Textarea,
        validators=[MaxLengthValidator(1000, message='Максимальная длина комментария — 1000 символов.')]
    )

    class Meta:
        model = Answer
        fields = ['description']
        help_texts = {'text': 'Введите ответ'}

    def clean_text(self):
        super(AnswerForm, self).clean()
        text = self.cleaned_data.get('text')
        return text

    def save(self):
        text = self.clean_text()
        ans = Answer.objects.create(answer=text, created_time=timezone.now())
        return ans


class EditProfileForm(forms.ModelForm):
    username = forms.CharField(
        required=True,
        widget=forms.TextInput(attrs={'class': 'form-control'}),
        validators=[RegexValidator(regex='^[0-9a-zA-Z]*$', message='Username can only contain letters and digits')]
    )
    email = forms.EmailField(required=True, widget=forms.EmailInput(attrs={'class': 'form-control'}))
    avatar = forms.ImageField(
        required=False,
        widget=forms.FileInput(attrs={'class': 'form-control', 'id': 'avatar-input'})
    )

    class Meta:
        model = User
        fields = ('username', 'email')

    def save(self, commit=True):
        user = super().save(commit=False)
        if commit:
            user.save()
            profile = user.profile
            profile.avatar_path = self.cleaned_data.get('avatar')
            profile.save()
        return user


class ChangePasswordForm(forms.Form):
    old_password = forms.CharField(
        required=True,
        widget=forms.PasswordInput(attrs={'class': 'form-control'}),
        label="Старый пароль"
    )
    new_password = forms.CharField(
        required=True,
        widget=forms.PasswordInput(attrs={'class': 'form-control'}),
        label="Новый пароль"
    )
    confirm_password = forms.CharField(
        required=True,
        widget=forms.PasswordInput(attrs={'class': 'form-control'}),
        label="Подтвердите новый пароль"
    )

    def __init__(self, user, *args, **kwargs):
        self.user = user
        super(ChangePasswordForm, self).__init__(*args, **kwargs)

    def clean(self):
        cleaned_data = super().clean()
        old_password = cleaned_data.get('old_password')
        new_password = cleaned_data.get('new_password')
        confirm_password = cleaned_data.get('confirm_password')

        if not self.user.check_password(old_password):
            self.add_error('old_password', 'Старый пароль введен неверно')

        if new_password and confirm_password and new_password != confirm_password:
            self.add_error('confirm_password', 'Пароли не совпадают')

        return cleaned_data
