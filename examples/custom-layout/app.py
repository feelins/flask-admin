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
app.config['DATABASE_FILE'] = r'sample_db.sqlite'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + app.config['DATABASE_FILE']
app.config['SQLALCHEMY_ECHO'] = True
db = SQLAlchemy(app)


# Models
class Information(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    lang = db.Column(db.Unicode(64))
    score = db.Column(db.Unicode(64))
    name = db.Column(db.Unicode(64))
    duration = db.Column(db.Unicode(24))
    update_date = db.Column(db.DateTime, default=pd.to_datetime("today"))

    def __unicode__(self):
        return self.name


class ScoreRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    lang = db.Column(db.Unicode(64))
    score = db.Column(db.Unicode(64))
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


class AnalyticsView(admin.BaseView):

    @admin.expose('/')
    def index(self):
        return self.render('analytics_index.html')

# Customized admin interface


class CustomView(ModelView):
    list_template = 'list.html'
    create_template = 'create.html'
    edit_template = 'edit.html'


class ScoreRecordAdmin(CustomView):
    can_export = True
    can_delete = False
    column_list = ('lang', 'score', 'update_date')
    column_searchable_list = ('lang', 'score', 'update_date')
    column_filters = ('lang', 'score', 'update_date')
    can_view_details = True
    column_labels = dict(
        lang='语言', score='得分', update_date='更新日期')
    column_editable_list = ['lang', 'score', 'update_date']


class InformationAdmin(CustomView):
    can_export = True
    can_delete = False
    # column_list不指定列，则会显示所有列
    column_list = ('lang', 'score', 'name', 'duration', 'update_date')
    column_searchable_list = ('name', 'lang')
    column_filters = ('name', 'lang')
    column_editable_list = ['lang', 'score', 'name', 'duration']
    can_view_details = True
    column_labels = dict(
        lang='语言', score='得分',
        name='姓名', duration='时长'
    )
    column_export_list = ('lang', 'score', 'name', 'duration', 'update_date')


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
admin.add_view(ScoreRecordAdmin(ScoreRecord, db.session))
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

    input_xls = r'examples\custom-layout\test.xlsx'
    column_names, file_infos = load_file(input_xls, 0, 0)

    for item in file_infos:
        info = Information()
        info.lang = item[column_names[0]]  #
        info.name = item[column_names[1]]  #
        info.score = str(item[column_names[2]])  #
        info.duration = item[column_names[3]]  #
        info.update_date = pd.to_datetime(str(item[column_names[4]]).replace('.0', ''), format='%Y%m%d', errors='coerce')
        db.session.add(info)

    record = ScoreRecord()
    db.session.add(record)

    db.session.commit()

    return


if __name__ == '__main__':

    # Build a sample db on the fly, if one does not exist yet.
    app_dir = op.realpath(os.path.dirname(__file__))
    database_path = op.join(app_dir, app.config['DATABASE_FILE'])
    # print(database_path)
    # if not os.path.exists(database_path):
    build_sample_db()
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
