import os
import os.path as op
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
import pandas as pd

import flask_admin as admin
from flask_admin.contrib.sqla import ModelView
import sqlite3


# Create application
app = Flask(__name__)
app.app_context().push()
# Create dummy secrey key so we can use sessions
app.config['SECRET_KEY'] = '123456790'

# Create in-memory database
app.config['DATABASE_FILE'] = r'E:\003_ProgramLanguage\flask-admin\examples\custom-layout\sample_db.sqlite'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + app.config['DATABASE_FILE']
app.config['SQLALCHEMY_ECHO'] = True
db = SQLAlchemy(app)


# Models
class Information(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    lang = db.Column(db.Unicode(64))
    gender = db.Column(db.Unicode(8))
    name = db.Column(db.Unicode(64))
    duration = db.Column(db.Unicode(24))
    engine_name_id = db.Column(db.Unicode(64), unique=True)
    update_date = db.Column(db.DateTime, default=pd.to_datetime("today"))

    def __unicode__(self):
        return self.name


class Evaluation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    lang = db.Column(db.Unicode(64))
    mos_type = db.Column(db.Unicode(64))
    update_date = db.Column(db.DateTime, default=pd.to_datetime("today"))
    # article = db.relationship("Information",backref='speaker_style',lazy='dynamic')
    """
    lazy: 指定sqlalchemy数据库什么时候加载数据
        select: 就是访问到属性的时候，就会全部加载该属性的数据
        joined: 对关联的两个表使用联接
        subquery: 与joined类似，但使用子子查询
        dynamic: 不加载记录，但提供加载记录的查询，也就是生成query对象
    """

    def __unicode__(self):
        return self.name


class Version(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    other_comment = db.Column(db.Text)
    update_time = db.Column(db.DateTime, default=pd.to_datetime("today"))

    def __unicode__(self):
        return self.name


class AnalyticsView(admin.BaseView):

    @admin.expose('/')
    def index(self):
        return self.render('analytics_index.html')

# Customized admin interface


class CustomView(ModelView):
    list_template = 'list.html'
    create_template = 'create.html'
    edit_template = 'edit.html'


class EvaluationAdmin(CustomView):
    can_export = True
    can_delete = False
    column_list = ('lang', 'mos_type', 'update_date')
    column_searchable_list = ('lang', 'mos_type', 'update_date')
    column_filters = ('lang', 'mos_type', 'update_date')
    can_view_details = True
    column_labels = dict(
        lang='语言', mos_type='分类', update_date='更新日期')
    column_editable_list = ['lang', 'mos_type', 'update_date']
    # column_export_list = ('engine_name', 'engine_id')


class InformationAdmin(CustomView):
    can_export = True
    can_delete = False
    # column_list不指定列，则会显示所有列
    column_list = ('lang', 'gender', 'name', 'duration', 'engine_name_id', 'update_date')
    column_searchable_list = ('name', 'engine_name_id', 'lang')
    column_filters = ('name', 'engine_name_id', 'lang')
    column_editable_list = ['lang', 'gender', 'name', 'duration', 'engine_name_id', 'update_date']
    can_view_details = True
    column_labels = dict(
        lang='语言', gender='性别',
        name='姓名', duration='时长',
        engine_name_id='代号'
    )
    column_export_list = ('lang', 'gender', 'name', 'duration', 'engine_name_id', 'update_date')
    # form_overrides = dict(engine_lang_name=CKEditorField)
    # form_ajax_refs = {'engine_lang_name': ajax.QueryAjaxModelLoader('mos_acoustic_name', db.session, Evaluation,filters=["id>10"])}


class VersionAdmin(CustomView):
    can_export = True
    can_delete = False
    # column_list = ('engine_lang_name', 'engine_lang_id', 'engine_name', 'engine_id', 'version_build_time', 'version_number',
    #                 'version_release_time', 'release_vcn', 'release_type', 'release_serviceid', 'release_type', 'internal_pack_info',
    #                 'client_use_info', 'other_comment')
    column_searchable_list = ('other_comment', 'update_time')
    column_filters = ('other_comment', 'update_time')
    can_view_details = True
    column_labels = dict(update_time='更新日期', other_comment='备注'
                         )
    column_editable_list = ['other_comment', 'update_time']

# Flask views


@app.route('/')
def index():
    return '<a href="/admin/">Click me to get to Admin!</a>'


# Create admin with custom base template
admin = admin.Admin(app, 'Example: Layout-BS3', base_template='layout.html', template_mode='bootstrap3')

# Add views
admin.add_view(InformationAdmin(Information, db.session))
admin.add_view(EvaluationAdmin(Evaluation, db.session))
admin.add_view(VersionAdmin(Version, db.session))
admin.add_view(AnalyticsView(name='Analytics', endpoint='analytics'))


def load_file(file_path, target_sheet_name, head_line):
    file_infos = []
    file = pd.read_excel(file_path, sheet_name=target_sheet_name, header=head_line)
    column_names = list(file.columns)
    print("File loading begin! The file is:", file_path)
    for i in range(file.shape[0]):
        tmp_dict = {}
        for j in range(len(column_names)):
            tmp_dict[column_names[j]] = file.iloc[i, j]
        file_infos.append(tmp_dict)
    return column_names, file_infos


def build_sample_db():
    """
    Populate a small db with some example entries.
    """

    db.drop_all()
    db.create_all()

    input_xls = r''
    column_names, file_infos = load_file(input_xls, 0, 0)

    for item in file_infos:
        info = Information()
        info.lang = item[column_names[2]]  #
        info.name = item[column_names[3]]  #
        info.gender = item[column_names[4]]  #
        info.duration = item[column_names[5]]  #
        db.session.add(info)

    input_mos_file = r''
    column_names, file_infos = load_file(input_mos_file, 3, 1)

    for i, item in enumerate(file_infos):
        mos = Evaluation()
        mos.lang = item[column_names[0]]  # 语种说明
        mos.mos_type = ""
        # pd.to_datetime(data['creatDate'], format='%Y%m%d', errors='coerce' 20221010替换为时间
        mos.mos_date = pd.to_datetime(str(item[column_names[6]]).replace('.0', ''), format='%Y%m%d', errors='coerce')
        mos.mos_sentence_num = str(item[column_names[7]]).replace('.0', '').replace('nan', '')
        mos.mos_person_num = item[column_names[8]]
        db.session.add(mos)

    version = Version()
    # version.lang = ''
    db.session.add(version)

    db.session.commit()

    return


if __name__ == '__main__':

    # Build a sample db on the fly, if one does not exist yet.
    app_dir = op.realpath(os.path.dirname(__file__))
    database_path = op.join(app_dir, app.config['DATABASE_FILE'])
    # print(database_path)
    # if not os.path.exists(database_path):
    # build_sample_db()
    # workbook = Workbook(r'D:\001_WorkToolsFlow\flask-admin\examples\custom-layout\output.xlsx')
    # worksheet = workbook.add_worksheet()
    # # 传入数据库路径，db.s3db或者test.sqlite
    # conn=sqlite3.connect(database_path)
    # c=conn.cursor()
    # mysel=c.execute("select * from information")
    # for i, row in enumerate(mysel):
    #     for j, value in enumerate(row):
    #         worksheet.write(i, j, value)
    # workbook.close()

    # Start app

    app.run(debug=True)
