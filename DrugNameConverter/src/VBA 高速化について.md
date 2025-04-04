VBA内データの扱い方の注意事項のまとめ (処理速度低下を回避するための個人的まとめ)
VBA
最終更新日 2021年01月29日
投稿日 2021年01月27日
はじめに
今更ですが、VBAでのプログラム開発をする機会があり、そこで苦労したことを綴っておこうと思いました。
筆者はJavaやC#での開発経験がありますが、その感覚で実装したところ、データ件数が少ないときには何も問題がなくても、件数が増えると予期せぬ速度低下を引き起こす事態に遭遇しました。
そういったケースでは、処理時間が件数に比例ではなく、べき乗的に増加します。
一方で、VBAならではの特性(罠？)をきちんと把握し、多少コードが冗長になったとしても、その特性を考慮した上でコードを書くと、処理速度の低下を回避できるため、結果として驚くほど処理速度が向上します。実際にいろいろとテストコードを書いて処理時間を計測してみると、VBAは想像以上に高速に動作するプログラムであることが分かります。遅いのはあくまでコードのせいです。
コードのせいとは言っても、コードを書いた人が悪いわけではありません。初心者にはハードルの高いトラップ、JavaやC#の世界からは想像できないような(言語仕様的な)トラップがあるため、そのトラップに陥りやすく、それが予期せぬ処理速度低下の原因となっているだけです。
この記事は、そんな罠を避け、予期せぬ処理速度低下を回避し、幸せなVBA開発をする一助になればと思い投稿します。

他のサイトに委ねる事項
処理速度低下の回避、すなわち高速化と題されたTipsは、Web上に多数存在しています。
概ね確立された手法であり、VBA初心者だった私も大いに参考にさせて頂きました。
ここではそれらには詳しく触れませんが、要旨だけ記載します。
◆処理中は画面更新を停止する
コードでの全ての処理が終わってから画面の更新を行うことで処理速度低下を回避します。
後述のRange.Valueのみで処理を賄える場合は必要のないコードになります。
一方で書式のコピー処理などが必要な場合は効果を発揮する場合があります。
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.EnableEvents = False
◆Rangeから配列への一括コピーとその逆
RangeオブジェクトのValueでは、二次元配列へのセルの値の一括コピーとその逆が可能です。1つ1つのセルにアクセスするより圧倒的に高速に処理されるため、ほぼ全ての値の転記(数個のセルでない限り)はこの手法を使うのがベストです。配列内では、セル毎にデータの型が引き継がれており、範囲が単一セルでない限り、配列は必ず2次元となり、インデックスは1から開始されます。
Dim t_sheet0 As Worksheet, t_ary2D as Variant
・・・中略・・・
t_ary2D = t_sheet0.Range("A1:E100").Value
・・・中略・・・
t_sheet0.Range("A1:E100").Value = t_ary2D
本編へ
お待たせしました。ここから本編となります。
第一章　参照型と値型、参照渡しと値渡し
この記事に於いて参照型と値型などの理解は必須となります。でも身構える必要はありません。他の言語でプログラム開発をされている方は概ねご存知の事項です。ですので、ここでは初心者向けに分かりやすくするため、少し乱暴な例えを用いて解説します。
実際のところ、この章の解説は、VBAの配列が値型であり、常に値渡しになる危険性をお伝えするために設けています。詳説は他にたくさんの良サイトがあります。ただ、理解に至るまでをひとつの記事内で完結させないと意味がないと思い、概念/イメージだけでもお伝えできれるように記載しています。
◆値型
オブジェクトではない、String, Integerなどのほとんどの型は値型です。コンピューター内で、複製/コピーされて利用されます。
Dim aaa&, bbb&
aaa = 100
bbb = aaa         '''ここでaaaの値の100がbbbにコピーされる
aaa = 200         '''ここでaaaの値は200になるがbbbの値は100のまま
乱暴な例えです。Aさんが、コピー機でコピーした資料をBさんに渡しました。元の資料にAさんが赤ペンで書き込みをしても、Bさんの資料はそのままですよね。複製されたものは、元のものとは別個のものだからです。
値型のこの挙動は、仕様/ルールです。正確には「bbb = aaa」で行われる代入の処理が「値渡し」であり、実体の複製/コピーになっています。C言語などでは値型でもポインタによる「参照渡し」などが可能です。
◆参照型
オブジェクト型は参照型です。WorksheetやRange、Collectionなどはオブジェクト型であり参照型になります。『Set YYY = XXX』で代入する必要があるものが、対象になります。
    Dim ccc As New Collection, ddd As Collection
    Call ccc.Add("テスト0")
    Set ddd = ccc
    Call ccc.Add("テスト1")
    Debug.Print ddd(1)
    Debug.Print ddd(2)
dddにcccを代入した後で、cccに操作を加えても、dddでも同様に書き換わっていることが確認できます。dddとcccの中身/実体が同じものだからです。
乱暴に例えます。Aさんは1台のエアコンと2台のリモコンを購入し、Bさんにリモコンを渡しました。Aさんがリモコンで設定温度を1℃上げて、Bさんが同様に1℃上げると、合計2℃上げたことになります。同一の実体を操作しているからです。参照型ではこのように、共通の実体を「操作≒参照」しています。
参照型の代入は、正確には「参照の値渡し」と言われたりするもので、厳密には「参照渡し」ではありません。VBAでは関数の引数におけるByRefが「参照渡し」になります。
◆参照渡しと値渡し(ByRefとByVal)
関数(ここではFunctionとSubの双方の意)の引数に於いて、ByRefは参照渡し、ByValは値渡し、となります。
変数の代入では、実質的にByValの処理が行われますが、値型と参照型によって、前述のような挙動の違いが出てきます。
一方で、関数のByRefでは値型でも参照型でも「参照渡し」が行われます。
乱暴な例えも限界に来ていますが敢えて続けます。
前述の「コピーした資料」や「リモコン」は変数に格納された「値型の値」「参照型の値」そのものです。関数のByValではこれらが渡されます。対して、関数のByRefでは「AさんBさんそのもの」が渡される形になります。この場合は「Bさんが持っている資料やリモコン」を変更することが出来ます。
下記のコードでは、参照渡しのtestHHHでは、元の変数gggの値が書き換わっていて、値渡しのtestIIIでは書き換わっていないこと、が確認できます。
testHHHは参照渡しで「Gさん(変数gggそのもの)」を、
testIIIでは値渡しで「リモコン(Collection)」を、
渡しているようなイメージが近いかと思います。
Sub testGGG()
    Dim ggg As Collection
    
    Set ggg = New Collection
    Call ggg.Add("テストGGG")
    Call testHHH(ggg)
    Debug.Print "testHHH --> " & ggg(1)
    
    Set ggg = New Collection
    Call ggg.Add("テストGGG-2")
    Call testIII(ggg)
    Debug.Print "testIII --> " & ggg(1)
    
    '出力結果
    'testHHH --> HHHで書き換わりました
    'testIII --> テストGGG-2
End Sub

Sub testHHH(ByRef hhh As Collection)
    Set hhh = New Collection
    Call hhh.Add("HHHで書き換わりました")
End Sub

Sub testIII(ByVal iii As Collection)
    Set iii = New Collection
    Call iii.Add("IIIで書き換わりました")
End Sub
第二章　VBAの罠たち
◆配列の罠
配列は大量データを扱う上で、最も高速に動作し、実用十分な処理速度を保有しています。その上、Range.Valueでの読み取りや書き込みが非常に高速です。サイズが可変でなかったり等、扱いにくい側面もありますが、大量データの扱いは、可能な限り配列で行うことが妥当といえます。
但し、一つだけ大きな問題があります。信じられないことに、VBAでは配列は値型として取り扱われます。つまり変数の代入時、複写/コピーが行われます。これはループ内処理の中で扱うことが多い配列にとって、大きなトラップとなります。
1万行×10列のRange.Valueを配列にコピーし、1万回のループ処理をしたとします。通常は何ら問題のないこの処理も、クラスモジュールを作成し、そのメンバ変数に配列を格納して、ループ内でそれを利用すると、1万行×10列のデータが1万回、メモリ内で複写/コピーの処理が行われます。メンバ変数へのアクセスは、変数への代入と同様に、値渡しの処理が行われるようです。結果として、急激に処理時間がべき乗的に増加します。
この問題の回避方法として、ループ直前に、配列格納用の変数を用意し、ループ内ではその変数を利用し、値渡しの処理が行われないようにする方法もあります。但しこれは、他の言語でコードを書いたことがある方にとっては、お世辞にも行儀のいいコードとは言えません。一度は無駄な複写/コピーがなされます。加えて、読み取りは出来ても書き込みは、「コピーされた資料」に対してとなるため、意味を成しません。メンバ変数もアクセス時に既に「コピーされた資料」となっているため、書き込みは同様に意味を成しません。
この性質はとても厄介で、オブジェクト指向的なコードの記述をとても困難にします。グローバル変数という手法が魅力的に思えるほどです。
事実上、プログラムの部品化に有効な唯一の手段は、関数の引数のByRefをうまく活用することです。「参照渡し」のため、無用な複写/コピーもされず、元データの更新も可能になります。
前述のとおり、配列は最も優秀なデータ格納庫ですが、コード設計を阻害します。その特性を理解しつつ、バランスを取って、コーディングするほかありません。コードが多少冗長になっても、無駄な値渡しを抑止し、速度低下を回避することが重要です。
◆Collectionの罠
Collectionを初めて知ると、可変長配列として魅力的に感じますが、残念ながらVBAの世界は甘くありません。特定要素を指定しての上書きが出来ないという致命的なデメリットがあります。詳しい解説は他の良サイトに譲りますが、Addメソッドの第2引数のKeyを省略すると、ユニークアクセスとしては位置インデックスのみでのアクセスになります。そして位置インデックスでのアクセスは、データ量に応じて加速度的に速度低下を引き起こします。1000番目を指定したデータのアクセスも、内部的には、順に、1番目からアクセスし、2番目の位置を知り、2番目にアクセスし、3番目の位置を・・・という処理が行われます。

14:41:50.371(経過:000.000;前回差:000.000) *** 計測タイマーが初期化されました
14:41:50.929(経過:000.559;前回差:000.559) *** Collectionの登録処理完了100001
14:41:51.144(経過:000.773;前回差:000.215) *** Collectionの読み込み完了100001
14:41:51.160(経過:000.789;前回差:000.016) *** Collectionの読み込み完了@ForEach100001
===========================================
14:41:51.179(経過:000.000;前回差:000.000) *** 計測タイマーが初期化されました
14:41:51.203(経過:000.023;前回差:000.023) *** Collectionの登録処理完了100001
14:42:17.203(経過:025.902;前回差:025.879) *** Collectionの読み込み完了100001
14:42:17.765(経過:025.918;前回差:000.016) *** Collectionの読み込み完了@ForEach100001
上記は、上はKey有りで、下はKeyなしで、10万件のデータを、{書き込み、インデックスまたはKeyで読み込み、ForEachで読み込み}した処理時間です。
Key有りの場合はおそらく内部的なハッシュ生成などがあり、Keyなしと比較すると書き込みに時間が掛かっています。Keyなしの場合のインデックス指定での読み込みは25秒の時間が掛かっています。これは件数次第で膨大に膨れ上がります。そしてForEachでの読み込みはいずれでも高速に動作します。
使い勝手のあまり良くないCollectionですが、それでも配列サイズが事前に分からないケースでは重宝する場合もあります。その場合は以下で上手く使い分けることをお勧めします。
Key有りで使う
インデックス（Keyはその代替）でのアクセスが必要なケース。
Keyなしで使う
先頭から順番に全データを処理するような用途（ForEachで処理可能）のみで十分なケース。
◆Dictionary vs 「Key有りのCollection」
DictionaryはScripting Runtimeの参照追加が必要ですが、キーのユニークが担保され、値の更新も出来るため、全体として「Key有りのCollection」より使いやすく優秀です。むしろDictionaryでいいケースがほとんどです。ただ1点、「Key有りのCollection」のほうが優れているケースがあります。
下記は10万件での処理結果です。少しCollectionの書き込みに時間が掛かる点が気になるくらいです。
14:36:02.171(経過:000.000;前回差:000.000) *** 計測タイマーが初期化されました
14:36:02.554(経過:000.543;前回差:000.543) *** Collectionの登録処理完了100001
14:36:02.769(経過:000.758;前回差:000.215) *** Collectionの読み込み完了100001
===========================================
14:36:02.789(経過:000.000;前回差:000.000) *** 計測タイマーが初期化されました
14:36:03.203(経過:000.293;前回差:000.293) *** Dictionaryの登録処理完了100001
14:36:03.351(経過:000.563;前回差:000.270) *** Dictionaryの読み込み完了100001
では100万件にするとどうでしょうか。Dictionaryの処理時間が急激に増加しています。
14:01:59.550(経過:000.000;前回差:000.000) *** 計測タイマーが初期化されました
14:02:06.335(経過:006.785;前回差:006.785) *** Collectionの登録処理完了1000001
14:02:08.855(経過:009.305;前回差:002.520) *** Collectionの読み込み完了1000001
===========================================
14:02:08.859(経過:000.000;前回差:000.000) *** 計測タイマーが初期化されました
14:02:50.640(経過:041.207;前回差:041.207) *** Dictionaryの登録処理完了1000001
14:03:30.261(経過:081.402;前回差:040.195) *** Dictionaryの読み込み完了1000001
500万件はDictionaryは諦めました。
15:59:54.160(経過:000.000;前回差:000.000) *** 計測タイマーが初期化されました
16:00:31.070(経過:036.910;前回差:036.910) *** Collectionの登録処理完了5000001
16:00:45.093(経過:050.934;前回差:014.023) *** Collectionの読み込み完了5000001
16:00:45.839(経過:051.680;前回差:000.746) *** Collectionの読み込み完了@ForEach5000001
概ね100万件から500万件では5倍+αの処理時間となっていて、爆発的には増えていません。一方でDictionaryは爆発的に増加しているので、大量データになる可能性があるケースでは注意が必要です。(この原因は筆者には解明できませんでした。)
まとめると、データ件数が事前には分からないが、膨らむ可能性もある場合は、「Key有りのCollection」が少し優位になります。10万件以下であれば、Dictionaryでも良さそうです。
◆全体的な速度比較とまとめ
最後に、ArrayList(.Net実装)も含めた速度比較結果を提示しておきます。
10万件での比較
===========================================
10:20:52.183(経過:000.000;前回差:000.000) *** 計測タイマーが初期化されました
10:20:52.757(経過:000.574;前回差:000.574) *** ArrayList登録処理完了100001
10:20:53.492(経過:001.309;前回差:000.734) *** ArrayList読み込み完了100001
===========================================
10:20:53.492(経過:000.000;前回差:000.000) *** 計測タイマーが初期化されました
10:20:53.500(経過:000.008;前回差:000.008) *** 純粋な配列の登録処理完了
10:20:53.507(経過:000.016;前回差:000.008) *** 純粋な配列の読み込み完了100000
===========================================
10:20:53.511(経過:000.000;前回差:000.000) *** 計測タイマーが初期化されました
10:20:54.054(経過:000.543;前回差:000.543) *** Collectionの登録処理完了100001
10:20:54.265(経過:000.754;前回差:000.211) *** Collectionの読み込み完了100001
10:20:54.281(経過:000.770;前回差:000.016) *** Collectionの読み込み完了@ForEach100001
===========================================
10:20:54.285(経過:000.000;前回差:000.000) *** 計測タイマーが初期化されました
10:20:54.566(経過:000.281;前回差:000.281) *** Dictionaryの登録処理完了100001
10:20:54.820(経過:000.535;前回差:000.254) *** Dictionaryの読み込み完了100001
100万件での比較
===========================================
10:28:26.171(経過:000.000;前回差:000.000) *** 計測タイマーが初期化されました
10:28:31.808(経過:005.637;前回差:005.637) *** ArrayList登録処理完了1000001
10:28:39.421(経過:012.902;前回差:007.266) *** ArrayList読み込み完了1000001
===========================================
10:28:39.421(経過:000.000;前回差:000.000) *** 計測タイマーが初期化されました
10:28:39.148(経過:000.074;前回差:000.074) *** 純粋な配列の登録処理完了
10:28:39.218(経過:000.145;前回差:000.070) *** 純粋な配列の読み込み完了1000000
===========================================
10:28:39.218(経過:000.000;前回差:000.000) *** 計測タイマーが初期化されました
10:28:45.789(経過:006.570;前回差:006.570) *** Collectionの登録処理完了1000001
10:28:48.265(経過:009.047;前回差:002.477) *** Collectionの読み込み完了1000001
10:28:48.410(経過:009.191;前回差:000.145) *** Collectionの読み込み完了@ForEach1000001
===========================================
10:28:48.414(経過:000.000;前回差:000.000) *** 計測タイマーが初期化されました
10:29:28.781(経過:040.367;前回差:040.367) *** Dictionaryの登録処理完了1000001
10:30:08.261(経過:079.848;前回差:039.480) *** Dictionaryの読み込み完了1000001
そもそも実際にテスト計測する至った発端は、Collectionが使いにくいことから、c#で使い慣れたArrayListを使うことを検討していた際、「COM利用なのでかなり遅い」といった記事をみたが、どれくらい遅いのかの情報が見つけられなかったことからでした。見えてきたのは、たしかにArrayListは速くはないし、遅いですが、Collectionも同様に遅いこと、そして配列に関しては、非常に高速に処理されること、でした。
『VBA内データは全て配列で処理することにしよう』と単純に考えられればいいのですが、やはり事前に配列のサイズが分かるケースばかりでもありませんので、可変長を使いたいケースもあります。また配列が値型として取り扱われる問題もあります。鋏や包丁が用途別に用意されているように、目的によって使い分けていく必要があります。それには、その特性も理解する必要があります。残念ながらVBAの環境は、他のプログラム言語と比較して、用意された鋏や包丁はあまり質の良いものではありません。ただ、それも理解した上で使いこなせば、予期せぬ処理速度の低下を回避することができます。この記事がその一助になれば幸いです。

テストコードと謝辞
時間計測では「Excel作業をVBAで効率化」さんの
https://vbabeginner.net/get-current-datetime-milliseconds/
のコードを利用させて頂きました。有用なコードをありがとうございます。
テスト検証で使用したコードを掲示して置きます。
標準モジュールに格納する.vbaOption Explicit


Private c_lastTime!, c_firstTime!

'''****************************
'''printTimeの内部保持時刻を初期化します。
'''****************************
Sub initTime()
    Debug.Print "==========================================="
    c_firstTime = Timer
    c_lastTime = c_firstTime
    printTime "計測タイマーが初期化されました"
End Sub

'''****************************
'''処理時間を表示するためのメソッドです。
'''
'''https://vbabeginner.netより取得し
'''前回との差の表示を追加カスタマイズ。
'''****************************
Sub printTime(t_msg$)
    Dim t                                        '// Timer値
    Dim tint                                     '// Timer値の整数部分
    Dim m                                        '// ミリ秒
    Dim ret                                      '// 戻り値
    Dim sHour
    Dim sMinute
    Dim sSecond
    
    '// Timer値を取得
    t = Timer
    
    '// Timer値の整数部分を取得
    tint = Int(t)
    
    '// 時分秒を取得
    sHour = Int(tint / (60 * 60))
    sMinute = Int((tint - (sHour * 60 * 60)) / 60)
    sSecond = tint - (sHour * 60 * 60 + sMinute * 60)
    
    '// Timer値の小数部分を取得
    m = t - tint
    
    '// hh:mm:ss.fffに整形
    ret = Format(sHour, "00")
    ret = ret & ":"
    ret = ret & Format(sMinute, "00")
    ret = ret & ":"
    ret = ret & Format(sSecond, "00")
    ret = ret & Format(Left(Right(CStr(m), Len(m) - 1), 4), ".000")
    
    Dim t_diff!, t_diffStr$, t_diff2!, t_diffStr2$
    If c_lastTime > 0 Then
        t_diff = t - c_lastTime
        t_diffStr = Format(t_diff, "000.000")
    End If
    If c_firstTime > 0 Then
        t_diff2 = t - c_firstTime
        t_diffStr2 = Format(t_diff2, "000.000")
    End If
    
    Debug.Print ret & "(経過:" & t_diffStr2 & ";前回差:" & t_diffStr & ") *** " & t_msg
    c_lastTime = t
End Sub
シートなどに格納してテストするコード.vbaSub aaa()
        
    Dim i&
    Dim t_tmp, t_tmp2
    Const NumberOfTests& = 1000000                'テスト回数
    '''''''''''''''''''
    Dim t_list As ArrayList
    Set t_list = New ArrayList
    
    Module1.initTime
    For i = 0 To NumberOfTests
        Call t_list.Add("sss")
    Next i
    Module1.printTime "ArrayList登録処理完了" & t_list.Count
    
    For i = 0 To NumberOfTests
        t_tmp2 = t_list(i)
    Next i
    Module1.printTime "ArrayList読み込み完了" & t_list.Count
    '''''''''''''''''''
    Dim t_arr_pure(NumberOfTests) As String
    
    Module1.initTime
    For i = 0 To NumberOfTests
        t_arr_pure(i) = "sss"
    Next i
    Module1.printTime "純粋な配列の登録処理完了"
    
    For i = 0 To NumberOfTests
        t_tmp = t_arr_pure(i)
    Next i
    Module1.printTime "純粋な配列の読み込み完了" & UBound(t_arr_pure)
    
    '''''''''''''''''''
    Dim t_collection As Collection
    Set t_collection = New Collection
    
    Module1.initTime
    For i = 0 To NumberOfTests
        Call t_collection.Add("sss", "" & i)
    Next i
   
    Module1.printTime "Collectionの登録処理完了" & t_collection.Count
    
    For i = 0 To NumberOfTests
        Call t_collection.Item("" & i)
    Next i
    Module1.printTime "Collectionの読み込み完了" & t_collection.Count
    
    For Each t_tmp In t_collection
        t_tmp2 = t_tmp
    Next t_tmp
    Module1.printTime "Collectionの読み込み完了@ForEach" & t_collection.Count
    
    '''''''''''''''''''
    Dim t_dic  As Dictionary
    Set t_dic = New Dictionary
    
    Module1.initTime
    For i = 0 To NumberOfTests
        Call t_dic.Add("A" & i, "sss")
    Next i
    
    Module1.printTime "Dictionaryの登録処理完了" & t_dic.Count
    Dim dicval
    For i = 0 To NumberOfTests
        dicval = t_dic("A" & i)
    Next i
    Module1.printTime "Dictionaryの読み込み完了" & t_dic.Count
    

End Sub