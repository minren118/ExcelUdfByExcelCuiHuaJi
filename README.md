# 开源说明
此项目为Excel催化剂插件里的一部分，主要是自定义函数篇，使用ExcelDna的框架开发，可使用.Net语言，开发自定义函数供Excel使用，使用体验也很不错，具体优点如下：
1. 可以有充足的注释说明供Excel用户调用时查看，且无论是在函数体书写还是在函数向导上都可很清晰地看到注释信息，详细到每个函数的参数都可设置注释信息。
![函数向导注释效果](https://upload-images.jianshu.io/upload_images/9936495-dbc5085ef3e6489d.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)
![书写函数体时的注释效果](https://upload-images.jianshu.io/upload_images/9936495-bd1f58a87f2e9c85.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

2. 可方便部署
只需生成打包成一个xll文件，即可发布给用户使用，用户可按xlam的加载项方式使用一次安装，日后长期使用或单次使用只需双击此xll加载自定义函数即可使用。

**亦可自行用程序来打包封装，实现用户一键安装，因xll文件区分32位和64位Excel运行，故需考虑用户Excel的位数将对应位数的xll安装到用户电脑内，
后期可开源Console控制台程序的方式安装xll**

![安装xll文件，效果同xlam安装方式一样](https://upload-images.jianshu.io/upload_images/9936495-edac214b83e366c3.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

# 和Excel催化剂插件关系
此自定义函数项目，可脱离Excel插件的安装方式运行，Excel催化剂插件仅是用代码的方式将xll文件加载到用户本地实现了自定义函数的功能，
单独加载xll即可完整实现所有自定义函数的功能。

# 问题讨论
因使用源代码的问题，与插件使用人群有较大的区分，现新建QQ群，专门用于开源代码的讨论学习。
QQ群：Excel催化剂开源讨论群，QQ群号：788145319

# 捐赠打赏
若代码对您有帮助，不妨以打赏的方式支持下，此开源代码是本人历经一年时间全时间开发的成果，有很大的商业价值和学习价值。

![支付宝捐赠通道](https://upload-images.jianshu.io/upload_images/9936495-a2f193dc1e62fa83.jpg?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

![微信捐赠通道](https://upload-images.jianshu.io/upload_images/9936495-4f9c7a855ddf9d34.jpg?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

# 关于Excel催化剂
**Excel催化剂**先是一微信公众号的名称，后来顺其名称，正式推出了Excel插件，插件将持续性地更新，更新的周期视本人的时间而定争取一周能够上线一个大功能模块。**Excel催化剂**插件承诺个人用户永久性免费使用！

**Excel催化剂**插件使用最新的布署技术，实现一次安装，日后所有更新自动更新完成，无需重复关注更新动态，手动下载安装包重新安装，**只需一次安装即可随时保持最新版本！**

**Excel催化剂**插件下载链接：https://pan.baidu.com/s/1Iz2_NZJ8v7C9eqhNjdnP3Q

取名**催化剂**，因Excel本身的强大，并非所有人能够立马享受到，大部分人还是在被Excel软件所虐的阶段，就是头脑里很清晰想达到的效果，而且高手们也已经实现出来，就是自己怎么弄都弄不出来，或者更糟的是还不知道Excel能够做什么而停留在不断地重复、机械、手工地在做着数据，耗费着无数的青春年华岁月。所以催生了是否可以作为一种媒介，让广大的Excel用户们可以瞬间点燃Excel的爆点，无需苦苦地挣扎地没日没夜的技巧学习、高级复杂函数的烧脑，最终走向了从入门到放弃的道路。

最后Excel功能强大，其实还需树立一个观点，不是所有事情都要交给Excel去完成，也不是所有事情Excel都是十分胜任的，外面的世界仍然是一个广阔的世界，Excel只是其中一枚耀眼的明星，还有其他更多同样精彩强大的技术、工具等。***Excel催化剂**也将借力这些其他技术，让Excel能够发挥更强大的爆发！

# 关于Excel催化剂作者
姓名：李伟坚，从事数据分析工作多年（BI方向），一名同样在路上的学习者。
服务过行业：零售特别是鞋服类的零售行业，电商（淘宝、天猫、京东、唯品会）

技术路线从一名普通用户，通过Excel软件的学习，从此走向数据世界，非科班IT专业人士。
历经重重难关，终于在数据的道路上达到技术平原期，学习众多的知识不再太吃力，同时也形成了自己的一套数据解决方案（数据采集、数据加工清洗、数据多维建模、数据报表展示等）。

擅长技术领域：Excel等Office家族软件、VBA&VSTO的二次开发、Sqlserver数据库技术、Sqlserver的商业智能BI技术、Powerbi技术、云服务器布署技术等等。

2018年开始职业生涯作了重大调整，从原来的正职工作，转为自由职业者，暂无固定收入，暂对前面道路不太明朗，苦重新回到正职工作，对**Excel催化剂**的运营和开发必定受到很大的影响（正职工作时间内不可能维护也不可能随便把工作时间内的成果公布于外，工作外的时间也十分有限，因已而立之年，家庭责任重大）。

和广大拥护者一同期盼：**Excel催化剂**一直能运行下去，我所惠及的群体们能够给予支持（**多留言鼓励下、转发下朋友圈推荐、小额打赏下和最重点的可以和所在公司及同行推荐推荐，让我的技术可以在贵司发挥价值，实现双赢（初步设想可以数据顾问的方式或一些小型项目开发的方式合作）。**

# 更多联系方式
QQ和微信同号：190262897
![联系作者](https://upload-images.jianshu.io/upload_images/9936495-1ec30fd89d19873e.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)


![公众号](https://upload-images.jianshu.io/upload_images/9936495-d54c58f58ae097a3.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

