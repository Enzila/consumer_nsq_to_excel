import nsq, logging, urlparse, json, xlsxwriter, ast

def nsq_run(topic,channel,tcp_address,handler):
    nsq.Reader(message_handler= handler,
               topic=topic,
               channel=channel,
               lookupd_poll_interval=5,
               nsqd_tcp_addresses=tcp_address,
               lookupd_request_timeout=30,
               max_in_flight=3)
    nsq.run()

def handlers(message):
    try:
        data_str = dict(urlparse.parse_qsl(message.body))['message']
        data    = json.loads(data_str)
        writer(data=data)
        with open("foo.txt", "r") as f:
            print f.readlines()[-1]
        logging.info('succes!')
        logging.warn('Succes!!')
        return True
    except ValueError, e:
        logging.warn(e)
        return True
    except Exception, e:
        return False


def writer(data = {u'ann_statements': [], u'pubdate': u'2020-01-31 08:07:59', u'image': u'https://cdn-03.independent.ie/incoming/a834c/38909724.ece/AUTOCROP/w620/AFP_1OJ3T0.jpg?token=-833551974', u'hashtags': [], u'images': [u'https://cdn-03.independent.ie/incoming/a834c/38909724.ece/AUTOCROP/w620/AFP_1OJ3T0.jpg?token=-833551974'], u'id': u'0b6e5d4d17412522bfc78fd021f3b6d3', u'media_tags': [], u'category': u'Sport', u'author': u'', u'clipper_id': 1439, u'content': u'<p>Garbine Muguruza will face Sofia Kenin in a surprise Australian Open final after both caused upsets at a scorching Melbourne Park on Thursday.</p><p>Kenin broke Australian hearts with a 7&ndash;6 (6) 7&ndash;5 victory over world number one Ashleigh Barty, ending hopes of a first home singles winner since 1978.</p><p>Muguruza then won the battle of the two&ndash;time grand slam champions 7&ndash;6 (8) 7&ndash;5 against fourth seed Simona Halep.</p><p>Muguruza is unseeded here having dropped well away from the heights that saw her beat Serena Williams to win the French Open in 2016 and Wimbledon a year later.</p><p>But the Spaniard&#39;s talents have never been in doubt and, back under the guidance of Conchita Martinez &ndash; who coached her to the Wimbledon title in a short&ndash;term arrangement &ndash; Muguruza has been rejuvenated.</p><p>This was a ding&ndash;dong battle, with Muguruza failing to serve out the first set but saving four set points, two at 5&ndash;6 and two more in the tie&ndash;break, before taking her third chance.</p><p>Halep, the 2018 finalist here and looking to add to her own French Open and Wimbledon titles, took her frustration out on her racket but broke the Muguruza serve in the second set and had a chance to serve it out at 5&ndash;3.</p><p>She could not take it, though, with Muguruza&#39;s defence a revelation as she won the final four games.</p><p>The first semi&ndash;final followed almost the same pattern. There was no doubt who Rod Laver Arena was rooting for but Barty was unable to take two set points in either set.</p><p>Since making the last eight here 12 months ago, the 23&ndash;year&ndash;old has won the French Open title and risen to the top of the world rankings, pushing expectations sky high.</p><p>But she insisted the pressure had not weighed heavily, saying&#58; &quot;Not at all. I&#39;ve been in a grand slam semi&ndash;final before. Yes, it&#39;s different at home.</p><p>&quot;I enjoyed the experience. I love being out there. I&#39;ve loved every minute of playing in Australia over the last month.&quot;</p><p>Barty conducted her press conference while bouncing her 12&ndash;week&ndash;old niece Olivia on her lap.</p><p>&quot;This is what life is all about,&quot; she said. &quot;It&#39;s amazing.&quot;</p><p>Barty kept perspective on her defeat, saying&#58; &quot;I think it was a match where I didn&#39;t feel super comfortable. I felt like my first plan wasn&#39;t working. I couldn&#39;t execute the way that I wanted. I tried to go to B and C.</p><p>&quot;It&#39;s disappointing. But it&#39;s been a hell of a summer. If you would have told me three weeks ago that we would have won a tournament in Adelaide, made the semi&ndash;finals of the Australian Open, I&#39;d take that absolutely every single day of the week.</p><p>&quot;But I put myself in a position to win the match today and just didn&#39;t play the biggest points well enough. I have to give credit where credit&#39;s due. Sofia came out and played aggressively on those points and deserved to win.&quot;</p><p>Kenin&#39;s talent and competitive nature marked her out from an early age so it was no surprise to see her rise to the occasion on the biggest day of her tennis life.</p><p>The Russian&ndash;born Floridian, who has stayed remarkably under the radar despite beating Williams at the French Open last year, will move into the top 10 whatever happens on Saturday.</p><p>&quot;I&#39;d like to first apologise to all of the Australian fans,&quot; said Kenin. &quot;I know they wanted her to win. It&#39;s not easy for them. I beat the world number one. I&#39;m so grateful and so happy.</p><p>&quot;I&#39;ve dreamed about this moment since I was five years old. I just feel like I&#39;ve always believed in myself. I&#39;ve worked hard. I&#39;ve pictured so many times being in the final, all the emotions, how it&#39;s going to lead up into the final.</p><p>&quot;I feel like I&#39;m doing good keeping my emotions. I feel like everything is just paying off. I see all the hard work I&#39;ve been putting is really showing now.&quot;</p>', u'source': u'independent', u'editor': u'', u'originalDate': u'2020-01-31 08:07:59', u'tags': u'', u'rawCategory': [], u'link': u'https://www.independent.ie/sport/other-sports/tennis/garbine-muguruza-through-to-face-sofia-kenin-in-surprise-australian-open-final-after-day-of-shocks-38909725.html?token=-1139038632', u'mobileCategory': u'', u'desc': u'Garbine Muguruza will face Sofia Kenin in a surprise Australian Open final after both caused upsets at a scorching Melbourne Park on Thursday.Kenin broke Australian hearts with a 7 ...', u'lang': u'', u'mobileSubCategory': u'', u'country': u'', u'title': u'Garbine Muguruza through to face Sofia Kenin in surprise Australian Open final after day of shocks', u'dir': 274394401}):
    with open("foo.txt", "a") as f:
        f.write("\n")
        for key, value in data.items():
            f.write("{}|;|".format(value))

def to_exce():
    with open("foo.txt", "r") as f:
        workbook = xlsxwriter.Workbook('raw_data.xlsx')
        worksheet = workbook.add_worksheet()        
        obj = []
        last = []
        for a in f.readlines(): obj.append(a)
        for b in obj: last.append(b.split('|;|')[0])
        # print type(last[0])

        row = 0
        for c in last:
            col = 0
            res = ast.literal_eval(c)
            for key, valu in res.items():
                worksheet.write(0,col,key)
                try:
                    worksheet.write(row, col, valu)
                except TypeError:
                    try:
                        worksheet.write(row,col,''.join(valu))
                    except TypeError:
                        worksheet.write(row,col,'None')
                col+=1
            row+=1
        workbook.close()



if __name__ == '__main__':
    # nsq_run(topic=getpath('nsq','topic1'),channel=getpath('nsq','channel1'),tcp_address=getpath('nsq','tcp_address1').split(','),handler=handlers)
    to_exce()
    cek = [None]


