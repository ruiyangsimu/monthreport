pyuic5 -o main_window_ui.py ui/main_window.ui
# 生成图片增加了自己的逻辑，不要随意覆盖
#pyuic5 -o gen_pic.py ui/gen_pic.ui
pyuic5 -o gen_success.py ui/gen_success.ui
pyrcc5 -o res_rc.py .\ui\resources\res.qrc