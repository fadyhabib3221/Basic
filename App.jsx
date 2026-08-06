import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// Perla Di Mare logo, embedded as a data URI so the app works as a single self-contained file.
const LOGO_DATA_URL = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAfQAAAEGCAMAAACkUpeqAAAAwFBMVEUAAAAGr/D3+vj4lyCs5fFZxuuj2Ofh7OpsyelOuuGV1Oh28vryrVdqs/kjtfsC/P4twvZdudsAAP8Aff/w0aQftrdtzexstbT46LDwypn6sWitr66qsfcTb7f/AAB/f/8Bf3/vtGn//wB/f334r6j/+XtnttQytOX+dxX+tyz/f3+1tW3X5eKI0unpsGfuyZIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACdJharAAAAMHRSTlMA/Q/9DttiVqPflgTrBQQB+qUBAmQDZQQMnAYDBQMBAgKmAQIJAmaiAgMCA43LZ9EIgK91AAAtzUlEQVR42u1diWLaOBAVU/nGNjeEI+Rskjbt/v/frUa3jWwMocRppd1tNwGMPU9zHyLEL7/88uvvXyCXp8S/sco8n8/VD8XdLis99H/12tO5i9Pn+dbT5i8V6fmE/xWm8UsQJEnC/hvHaUhX7LfT+d5T6K/T4Rlja4AwHjjWOI0Yz+fU0+nvw92NuFgviDssvXr/a+Auc6a+08GRNQ4Z5HfEm/R/A4Nv2H9RPOiwEgb7BNWAJ9vXX90gZ2vEYAev2r8+ozPBHgy6rzgiU8/pX3jdIubRKZDjSgGWnnZfdxUE4sHJKwD2Qb++qGSnp7O5WBH7qJfxX3LtSHge5ijiPfm+JKMXJBycu0apN+K/qNU+Gpy/UvBG/BdcH8IcUZ94Gn41Rk8HH1wLH5v7Yr4afJDPcb1D5kH/OnxOyQUwHwQRoR71LwI5ZlEHl1gJ5F7Cf5E1OTcm4zDmvOP2ZTj9QphjaK70BP0Ssv2ceHujgL/z4v1LwB6OLgb6ICRTT9Kv4KIPLrgS3xHxBSCfXE6hq8CcR73nK7+Qt2axui+I77toh9MYPY7j4KgBn3nC9npl3UPuCe9w4FuFtlXRjUnu6dpzRu8aYsUi92yym+4maJ5DS7mFD8r1G3Ta0UXHAncy0anTDX62MYpHPej9DsxE3TCPgOaI5LKcF9M9iF63KHD792Mfi+03o//qVtrO25TzHIxawMxcgz2Q+Lx6r1HvxOgxD7eUXJMvSLGI8OccBUVDUZ0Pz/R4lZ00esAEwgb52jSxJqL4lYn4X96S+2J8vodOwRYhzat2G+9hg0fq9PKXxMdneuujd8q0cGFO4eCtQcT4uSSRM+niPfUvLd1j7D+nru3BRMAGCherh+BB729kpoN0x/QJtQy2JA5pKAJyY7YdMpct991zel8X7ZJqiRl+zIYbWaocjTR4lxzt3Dke9K8djYMSZ00E2nkjGQWaZ1MSjURCzVV2s/Cg9zccdxzzMUxwFslIY55zvUDInKTcNysLEntD7susaZd2xRA2REsEXuAsTUCxZUKydIPuk6v9XDvSRboDuuOqJkZUUfKY7JZ76N8R9NRnXL6QwxZ0CcyQpUI1YfCTAuIRuuh3U+T/0A26j8h9aZXOlLOW32PmfYM05GPRE4N23SHovjaytw5bF5WOUbupBp3tgL3eKigmEvYiPdQSPrXa19WpTgpRnpDAgP5YkQ9MuhOH4xcSD3pPjfcuBZECdM3ARbUDKoUlgQ04VLrPt/RzFV2Md8rEuwY9YVwNVrw+xW70OVkcqnTf49JTOy7rxunocI+0085dtkAXSnJ/LvChmS8Deqdsy5xZalqno89GGKtDSMgCB/3zGiqXw+ZDM/1ce0IHXaz3OzuIk0KhtgzJMDpHyWFJbAx3nrx9Nd67dDMx54vhbLR2RLZANrSgvJaGiQGHjoi8l95b0DtVzWAFRWbJhIQhOt+IK2z2zpaoGGhG55mX8D1ceafpkBiGrc6p4BWRBSlQzle6XJKAr0EcCkafFh73vi2Hq+XOskEt/JIs+EbA2UTSjI8XGayEqhch2SRe8CFTT3TuCd0vTu8EOkbhoF4dn7zE8UsSiNop2dQ4Yf9S643sPSFqd+ony3050BN00Zpm0ojGl2mWLblZd+CzJ0GKpbTemP9aOh0jLXfonTcM+y4oUccyAhSuaH4Qgm9y+mKgJ3xEWHYYso0jAllevWI0cre1e9C/lMsWcDYuyYHg5oK9tON7GWlqXU78EJqeLNoF9BjIRJrlFS5mv99Z18Ig/KS5F/LFH/DSkwUdwrCpPnJtVdkj/HQmqIr2lux87EHvC6cfTbgkIciRzsCP9NGop1Ad9YwvtqXsUg96Xzg9g2MmXC2aI9tWR6El2m+ZnzafR+2n96U+1dof8701n54+AJjMCYgV83xKTvfy98spnzl05MBGD3pvVmsFdMIrJKoyXBS+JiK0XpZ0wo13iI7OlfOg92YVLSG5IILDExpgLrOxwSIEeTJAGncov/Gg98hnoy1WOw+rVjm9JGee7ONB75HPBk2i3fXu5/NP/PCg9wl2l2gexRhEd9l97cGcl5c4TcOQCfyXcVLP1HmXrTfmuzOeYgIylf1BSfPwsSAOa5oAQlvXe9D7JN8PcRxH4Kh7sFvUD8JtqSqUKbJ8mecZr6nhZ/sFHvQeog6H2hecVY3zhjmwIoG2y+oROpwuyf36wIPeO9Dj+oywqetdTTPhA45107npvLMJS6pGfhR4r1ZYK4q4Iy4E3ZhzRZ5njbqDfeyxQF8+9vn0XnG6BWYQAWxc2Kycp7bh2zscnQ3Y8+gR7xPomQm3pKLE1eXNO/g8iKDzvvIVsf1aesIA51s3jtkh5jz8nnXvY/Gc3i+njZtyQcoz4s530EPZzoRC7kX2V9bqzP1OGmU1sn/dP8eCSI/411p5xaXGEIpsOndhzoz52qnqwTv3wCsSG54ff4j1/MNvhz5yNqkKcixodctqgENHPmF6YGfp6Ofn9dvtW83Yv332yPcMdOZc38HEPu26YVQIzGk9DIcNLUq2w/PP11stLtZiraQ5+Lp+9rTukeGWBOicTYitmMHpzdXrW3GmPxWS/Xkt37W6uSFDtr6Jhf9LblZ4vVsCXvf3Y92BajQCSpt3BkztfImOsk/4ixitQe6e3Wuwa2s4nCHcPz3o/QCdaN2MP+VZHRfIaDFFdq7VOrIPFGyXwMMb3xTNgCvc799RIniK9wd0ye78V8WuyPOM/ZPnpRooUz9R8xeQkm0FLrchum/HW+PO2H298jTvEehimMCBq8YsvcXYEXMt4PmNdGHxCuzsgz890XsFOgc0GZNFKNaCxIm7B51sXjmP/+4OuECd6YJXT/VPt96D02oakcspkCe0xIcnQi6Z/c2bc58LeganIJ7wSRN8SOCMfDtvCUPer09cJU7oP+W8dL5VZjXTbXh/fzMjs9mMkNnNzREtz1D3rtvnLkyILr4fl/EBjpAiGyyevBlWXfDoAEMmCH43I+95vSdiniEfB03QB8GCaXJ4xDfaiA+Hv1cC8Nu39Xr9hOvn+u2NPGK0JpqRpljNjPzwNP/UNYEsm1KOfBQuYvKSVEeFCTfuUfpnFuL3PAXz/PYTHurXfHhac83fYN0PS0/2z7bgcT1mOs8CzxBFzGMzI11/Yp37rMLkhOfc31q8blhz9e+CfegLIz93PZCfyj4DmuVFUSlJz5YboaMrohr1OHl9e4L2yid4Wq2csA8j4iNzn83pq/XKAqEsgeKCp7c3gXjFKBveiLNVu/AqkxW3VTtAcboH/RMRL2c3M2mLvb69iSz4G3l9fW3yzm5ESq37N6zZu4ce9B6tFZe+9zeHXhRj8NmQOH3sk9xs/vYas3vQPxt0icO3+/v7379nbP3+/fvm3hFhGd7jpKHXMyqYmV0QDb0h1xvQyaxr+PQGgzCrc+C6XcEzqYh4Sh487fsO+vB+xsX07bmxH4JRHHO98twr+XUt0IVT/kGY1sRC3Qdn+g36kKCLdruGj6ZJ4NlIeA/6Z641uWlFHM11slo9XcZoVKj7MOwngx4N2xQ5cvcFC5wk6h70zw3OAES/74c1/2w4vL9HxmRi/fmS7SkrEFvMg/7J60EGYphrjj/e39/c/JahmtvbC2dAYQVcmwz9WV2fvH68vToCLq/r5z/hSgtjznN6L6T8Aw+9v76yP97I+ufjn1Mn6C0MqSf5v7W/kNU9Gf6twAAAj70/elL8S+4CufdFFP/cumUCfkWePCH+KV4nM/KbrD0h/i0fkcxWvmn5X2P1f6BxFdTy1SKKIM9/bzadoZwt54WuJ9/Ni63v4vrbV1FldfGTDzv/tesRKOVHVpE4CeR58eOxOMmKUs/uf+PKcRpTmCbOQ+0Yzy897H+doYLsHLf3eXvU/yrjjUv14OiR1N6k+4sEewaRNdwhScZsJYkDdrL7GrBDuZ3viul8Pp/udtOy9Ju1sp6BUDNlLxgvND/XDyaTxxpmPScgZJkj8+3t0IpkX2rB/hKrmelFsZPHV0BUnboY255cD5/nIaNiVgFdkDhmAiuOCW9ax2GR/Rj0Cp/PNqAmY4854rv5ptScvskytOijuHZUcY93Mc70pouDU5ITcbyiP4hFDM0lYjI2n45cOHTfFk8jsyemJ/09FIEfouzyOfVpXJ8upqAa9fqkW+CHmiXvCH+jUbQjlaHpYU9B32iplcQphbygYVzh+SCNPlvC42DqwRjI9tNugFlwv7SibqMGn6hcQb2HsBfyHpPUttogs40SnP392dpUHe3+aYImUKeZHVs0J8aPH4W9i9PgSU1CGv3ChvTpMtuXZbncTfhrlsjn0UV4/Kz7p1y0fp5htBdUSiKcht9FLFgiPiLLnqE+Ie8jOQy8yKwSddjPEfZFcPoZfH/mNrkCisnd57CGmKub4KzcLqCDPYg3gTzrF6NDyKNLjIUcQ0lK9ojmxCaRSvgU5EEeHJh8jnyXsp1h3tmPySzUe3bGMHsajmmIPlvTe2I7opx9yvzHUplGIXxCPTVMuA2XROUWOm9Sm9fDXqFeKuukcB+1ynU+mOOXg0+yozQBI5xpc30jLpWnk55oLI20gO9VZFPY7Wwrt+1zyyoJPkW+Mgpq9ai26/XUZA5RV7u9ZnxaJxhv+4I4FU/DpA/ctm/16Jd1Vtf1DehMqZhYHSQnThq7pkIfn2bOMIMIci3g0RbpRSMnWqK/lPCBdv0EZtMy1PdXF+5QjW+J0Vfba8gceKB8xyWQnSrjINtpAZ/2xW0DScsx3G2OPA/btAb19j3yJ+QrCQcV450i810p1FUS7t+k53zZVlMt6Ut9tKJl2KHDEDKifbfg2ptzqsRkrMcZh5yO17HiArXdTirmLtGUo1rA05J+6Daygq0LPPFUBDwGcLxql1eFpSZPfFVjVDnpuD2XUtwHPLbw5+/iWQYywhNNCAzVA3rrofbVJx/aemoHwodpGZwgecAYo6P0qkGauaYc8HQm+5P/YnGVmCxIs+fkB45EOYIx5T5wDyWJ4peX+OMKTTHQC+y67TYTbUiiK5olpZbuY227B8IW2f3xL18qRs9OJG3IeaPcERXa+gBgoOKi4Ydb/ZVK7yh4AE/eHQwGVw+HUqJsYCnddbToCrU0c+EsApwE+kScixVCRsuwoprO81gVu308tZwLs5TZpfnJ5B9cUcDPSVplFpnkRL30pyM0QCIhU04Jo4IMK6BEL7V8T8n0/OhQfClWy8lCmWUdbZPMctevVrHGNHiig5nCX9Mq/o/vu0JI5/AERjdBjQj2pMyUbgrOtuS0D3COZXHwQIuqpmxX6BvuOGi1nh7S4Y8gANpJH3wXX7k5UPF/bk2ldD9Fl6o4Z8pRLpRST84335VaXZDdx3exAP2lU456SbHGP9UIDCIHyJs/4MrdgTGF8CtLQrXA/NO5KyqF8ykhVLhTTtEc+CUWHwVdpb+Z6PhwilGJ9+OaQiQK0YAMQbP6+PAoRkrIxUvpbKeHgiXrGBR38MdBD4VIOaEceClNkBB4E2Ou9uj5oKvMQ3KBKhYtN4/jhCmXRJpS75rVq5kj1GVRBJcHfaGkOyc9SOF5ldIEyRfjE+p1tGMrAymZut/zQZ+QQLJZcQlXqJMzAUSbJmPQhbMYg59W7S32nlFw6TI6zejaYVOMHv357Krkse8ncDo1pJKX+CjoOiBJLpC0YRcbSRbaHSO8jLunQC23zeZqJtpl/OCyrD41FRw8HEc1o18jhaGM9+4yhdEkkVtUsZak8tmgZ0YiP14CdOUJlN3eyL52C1SzelzZLPLZ4sua1IXy19iF55bLihuw/OMHAhXwchroIhRnG0qaSc522VRQL7nkPj4aIdS2FH+SrcVrE9vCVJq3uCDVqfEWuK9Mde6FmRR/XLqr3X4S6Er/lvoJRtIEOY8wpbpkfBEbRkc5johkLV/GXIsbrRqasI5W9d8vaV5ZprrYmbqE5iqJ1dNBz5TA0+EcTbxzCaNVOr1IeaDFwa2Z2qJc2EWdoLkvti1MJTUuGBq1fYXkYYLPH+n07hUKTFUwsHvYPFek0gYHJYvB4NSoXtWYXFzShjGWeHtmujSAQs1z1p9bkvDjGUTHHQamtHAORp6IW76CIfdymvWuSGU0X14l3hmmrL5keZkqO8U447ZIupYvyv7cadJHoCXORLeAXI7mpVV8LSr1IvNzeYVaaGn1dA7OqDaCwXel+Epixc3PgkhFncMLiVCjMtuEcqZiSrHMX5tyEEvZMeEbLiIgl6z5nFutQSvmsE2IKd4hfz6XjqCLEvHOoBMdMFVMtNeWHT1z40cfzsfXb1KzOm0BfVE1Tox8t1MeeDQ70Ev2n5jNxVinnFvfi9vtGoMSlioi1/HL7nQOtKzbTecyqu4iv5wuK0B7bftm6a7eUz5W3FehwXW1YFZcthKdWZfBwA7HZVaDFVylpFg1WXRMuKh2ZiaHiprPebYZpsIi8cVcYdSSIx01hPYnMZUEUEt8mbdeNgYajiyQH5m0N43T12kdMK1+nbbYVlHTBEyVHVdLA2d0nhdFzZDi5HvgY0EolaNGQcV2mm5hApRdaf54mgDTnYGbJqUCyn7O65GYeoTkopb7Bv2lXwOda8iNRm8qj0MK5MUc3BJLcy9bdbmGwb0yQ/LNcziQhGHZSYjpQIxhhYk2w3LbVDCWkNpcNLMN9sPolDulvz+L/XVNeaP8ybWjqYMDmWotY86M2X9waceZ0XscDrRw4xInOPxevTvzRs0LOaqe1i1Z0onTkpO51SYpuHF7VwSqosI4t8AHtwFdpHG6gIqABOEQy5Yi4K1mGUmrET772zOxh3N2KUpoVuRdc80wjdqL2acm8H4Q1ZGxRfbzPpMvXFCy4nfGOmZI8fGTVN3LXf3tuaAAEpP9VdDdnSyoe6SFxelT58giSW+OBLX8V5C9m+4Hg6pGMzFY5VroSEgqDWXsjbImPQT6ZiBaxEGQyhaygA/6EaBKSVFNq4KYaQdUzYjh1Q1diQ9GVmOb0GHzOUxgMKgXB05/vBjQcdTchF9nEXaOmMB2mWPPRp67WsRAksvK9AjTIoqdYolBjhSQ05FGqSBLKbaBnOYYBFBivVcUMN4bJdpwxStTe+Aj+7h2ZUppPIYu7wazK/bNZyqgsYBljdGlC8dIWB00F2NNgvnuBOPcKqUZsQs+qkse0hV95MCaEhLG4ySFW+jKUmlLYYZWKsTSSjpVo7II7CsT2c7b4Vuz+aQW5wXn1yagN5y4y0SzUk0qTdkdJKYbXFBgpQyQRAWvM911XfOAoDLq0XoIGQ9y2RBY9DpKqNG19KAqRalGFb+uDpkTrjJY3x2jBSXtAmwZtI3/GmiZPZM4iLQxse8KukTd2XvuDCNWQAdrOmYCR8fJPs7FEKMwjeNxnIaUSwfbjuGyg/O1TlbAkseRQlAqXYMuBCIJg5FpsEy1OKhMcNwwY0QPN5CsJ9RnWJ2cZ7qxsbM3aLB4hDS3QrRFjVTAJHAi7ggVxrOa9RKE2heN7ZQSXuxOBx25uV6qCZVVdWa1EadhEIcqiX9KJk/5gg7UtXSvwDk3oO8qd520Z9OBGx9McgfVsZQIgS0/nyh+ARMbC8WjuWhZpIogUzuuqIcmhGHwK9TiAfW3mZiD3ra+14qpxQX+IKbvujrEBFJU6Dc8NFOFNE9VKZUeOmYirjKgN4qkyWUG0aWmfcC+w1DrcEyCEbLZ6JKMvBamC4xOZupJaejbXXeTSdkb8cGGzkqt8Ze2UbvQdWpWPPx4p9G0LkmtOXW2dsSLJrA1Rf5zfPpIOW1gsoxmJhrXTxYPhGRPwRqDOWZKflGf48h48D1R8zyDwWGYHFRS4cDMFZw+toxvvdsyYvvYKVAweyJg+/s51nuPMVwQq7u0+J7DvFS7vKJwuGBT3LDLmNNvoun7EyzlXPThDMYRzWu61VVooWsqUZ5EQdJpoo54cAnjeBHiWoy1IRsS27Riv8ZGX63SEe5A42lJHdNXGURkmzMCSKATMYw/jk2pjS6elioyI8In4OK+mOhelApX7ymHLj3oKAI+sCexDuxR/CGM9yxXtdBgxbgSe9xYSnKw8kgvdjNJaenWpJLIM+8S065K075/aupF6XU7zowb0RWcr4BuK8WoLeFSMOWvRmJbGy6MtZ2lkxQpvqkkWoXvBKMrDWepLk0BOe8LlDTgjVd8yQDP1sxsVRXz+rHZhrAmH1SijLJynam3rE4zLjh0UbCuI4v4/tBNwRGPvSgVxG9KT9CZisCM3Ke/iI5BxpyQKrSTloUjqyPl1aQEI+vJSbOQ1KWSyvEjug4xzLZu0OcVv72lcZ3pNRrIMYXswgWPjGUFNuoIY0sjwdNACRLK5KxSYem+yE1zZ9xrqRMTIdoyo99xm/KhGcqytapwJnthAstPp5zyQT0IpTVAwG+vnn5mrM7ZpIBbhGgKxo4zFcRM52RmyCi7q62JYlujsuTWG1WKw3RSO7TVGdRaIq1O7pNrDEAxQij8VskJzhyP6n0UZZCPmdFmtDkuzW8WjzuYFpWQxpJdgn11IBQ1SEHIFMVcqeAVZwUzeoao8FRFxmwFQIbx1ZvkbjbDkmSJA6gkY8Kj3SoLUddQAHcwqluPMpqZ8t1QVEzeF/Z7xqOSchGGELQ84t1oqhwoVf2YRSlAR4My0V7fg1UHW/Eh9VgOFQ3QAZq7kwNgWuEyT1/PytBx96UTdFE9bXZkY4ZWWg0MWUdlFo6hNhloHnlEdaJcoDEnWcBuASqWA9JSIjVivwL7hk1diC5H/WXNQMxlVMtkvwvZUXDoweAv+CWiWhxPTsTAYA57Yam+J8byzZW02sRjaYEZkUc9c8qMGJirrOVgsIhGxleaa4cttzagaoiEHCr6baBpcCLywipEW3pL+aO4w/25sVmWNl0byq4wncDv/xdktOEdhG4FGVVYBr1czREBZxm11ZS0MyI5BiqD1nG9aDM3ijLRMnZiN0j9Yv+nyr2cLXqKOaV3cJCBjHHof36naFLo4aoM8zujIHA/LK2oiLV5Er0dFxpR/KKg1rQHmenbplAp1jmzIIsbmVIMpZE9J4fPLXBZ7yuQZRbtSRslCFtuS+ebuPHyzt63MYyNwTgcqK3EvQx+7MG4PhupQw5aIYzKS5RhTwE7Le/MjjZVrJEjbcQVfirG+0+dOjHh51+YkHEaaA8YobszriaoEmRLopiZOolq4uLmeq4LMHQDWalbTVLpKlr1wecV1jPXjSoLOUm1WMPn4NfL9lD10xMxEi/XVnWDwzbhNIujDrfAMUfznmlZ2wILMaUSD+yahkzzsJIwZlrfWLsCOpA00GEbfh/wqNszlmrUQ9MwHfRp30fquDVLv+I2EwMIgu/K+NeHOb3ztle7tgKWytAMrMKy0io/0ylN3hWo9go1TlZgalLxPjZ3uuHo+/lFdIU5jYLZ2VqnFIW8fyLOdFlYhfDabm0q7RZTUrvMURCWYyLtXyl1SCwafxV8L4IbYWuEOz8BDcwuYOSjOnql6bSFSnJLZ+w2coqiCKM02rmJTp0US2NKVk/v0KP9cX9o80V+EwVVJxmAlcrTVdI4oFA/AZq2cU0RUF0SrrTX1ATz6dmgcwHE/OmkGouWj00lUZSJJVR6qXRS00gUKn0e2qGANeS3j5fdaYwViFDp3tSGOh/5YFcs2LEkex6ZNSkpCUIdgdhLRTl6b7ZDTURtHILKwAvPjGLmrIY7vkmNjbDCsyp2FlecK+MsgtXBg2DXm4ZyY6iLB8yRnYKPF59TMcg+tNJBQcBPGeT3wWNhEvSREKuZHsbu3GsgmhCTDkPQqDq0Aymi+DKOdUNVYAcHqKEAFalt/J567WkOoTVq3zzUKA0NpcMOp63wSXqJvKNQim1+5hFP4eqkUxKMM9GYru5ARcaZC8RvmSmRynEKxO4as45/UW67ZmHTPErlgWpopQTBJYro5LBd1xGy0hTLhY8hkzBTPabK6SeKcGWHYT082a3VwF5p4EC6jFZse25HPYXVimKWfY+kQKpBn5tJ8RuwJHGgc6w7eQhTdKT1kocvqslwJU02O2v6ESeJFcjJrUi2GLRe6eJnVgw157MapRZHcOilK8N1pVPC7GOjj06vUrBnlOrHSKzTgvkNqAylbC1SeiV0wyof4/vRzn5hTikxdUcyqw6Szx1RLhwHfaopINUjRyTVlp32gTXJqYlsBCYXJyN2HY6HEZaAdcCJJaO3oPv9mHuaEUc9TSIHjobVYgCYgH3TOF5/QRahcT8NwmDF7HWOIokuOHWJqkgWetlZGIrEN9jFNoKy+qncXrq61aM7Ub4xlPOGtYPGN8zWAp03LWdGVwOXdXzGQhJV9oElO8cq1RUyilLdUszYK5VcS7sWnUSLl5H0bzShdS49rCe6tlp68w1nHRWwL3n0Y0ecUc87U3y1rNejiv4CoeDC0MTxP1yjlumBdbZ/LvdDajmGqgOmwZJQ1bThsYYngbkWzFOjgBO1uaw5JFO9J8AKVUV133FpwkgVMHR/3ns8qs7KumtxdwGegd4ZrgdXJSHUq1Bo/YhtqJRDVvYy0ou5Sjnd62kGmlvmRrZxGu0wc5BUDJ2PrSdFFZyXk2VZTpdLPR9cF449EqtgM3aPRMmewy6gLzOp+RX7TIk5+IbrDWqDbgXjuCH1Q+ZkgprvqOP34vAKgD2SNAPtyEnuyxRYk+Ol3CDbOlzG2BjmTdXEIvtrlWJDnFCUUrrUaVmVEGHNFdMUEVnGZ66rQhJc6pgYnYtIHVBJi0rVFaikTAOqO10n1uxGTlTdaMB8abAUvFR2241g2pEGfaL124LviIyHAqBeRqjBrR0+Y5+qw8/p0IUQqzRIO7SdQXUWSFaGTeaUVa4fR1amjpfUp6QoGoZ8ZVYYTyn/xModPE6lBXRai3iGLIyLx9WqMeVMnejjdMJCq2YJmgs2K9pu0SZ/CnnumZlYZXS20hsV0C2Vzvcar6mLle0HOqprQrkVMDaWyx4S+3w9GJ1xzA7Q5qEeWtvjMWhLi9EjMaJtqquWwP05RVedyxwQwV4pv+ZBf0WbhCqrwphJrAxq4RSGOaPO4+Emj63otinYnDZUSAXtlgZgXlt8HzV6MrcHj9RBz00zCZN+e9ywIpAX1tKquaZJlShbpTCFBWfC4BiaO72rWI8XGsOyadObangZ2cE7xrBiQ4ZSf049S6bdmZCXduCmMWWEK3JkhEw2EQcpM4WS59g1JSOspMjyDExVtCjuapDukRwepsyLpuZhbZC5T5UWHoEcTWyxowmZJ1LaZPZYm5yU2o6j8qWQGDu4rFIuqYG+0aBHZW64T6RAwzPCmFFjbnlqgqyV34o7zvemUqS+WYjia1J9/gF9zKh0b3VYsp3T9+LxQjK2DlJNxiRVAalJThTmTpiEKFNmm57X22SoWf6027xnbL6Q2Q5LrBo7ToZv2R6zxHtuG+9THqmKmaYLlGdb2lWNNS4j1ji1QWr/thCUvDs1iG3VmsBhDXl4cKq2rJ76Dns9HPFAEMopGCKHKqZx6iG5Sh1G5mwdPvhr0pDaxpq2Rew+Kz0J4u/m4A7pXcAB5ni/AWjprjJxDZaE0WkhcZz4QydW5w7YezOohdErLpsxj5CUIpDHDGGrIYbOl0bspodi14piK9rseA73eOm+I02kdMWh3fKoZgDwLAun0VZEt3gEPoOGKKrx/GE7pbzMJDQXioR7CzprG2Ixuvu8D14eEwzaVhCHsc4iOktgArvWQFO/qVbHUs7oGRk7mSd3clOxXCuCPhh8UFZBX5hgaBSoeKAuc0P5BxNqIpllg/zhNZE8gk4L3kRzxgDWzFRgOFJK2oXnLsRkMpmKzM8v6eurErB5A+jJuwj7Byk1B9qrZKGOQ2D9ZRS4+lWwGEA2zYxfxkEb+hEvaFgy0/6hsmmmCPLIhBKVT9SYwp9swQ5N7Cy64GPaVfsVOtJBbYSDpk9aAX0gapxxEJFWCShmwyAGMx2/fgZZXlp2onS6kTbJGdNtrTSaS7HurT57cYojT2qgYwSEOgLstZ00UFmkCEYV7kSpvzXv4nA6pkrmuvpT3e2iie+TWMKQUewy5GuLOiOseHJWVGLXSJPYylHwvAUzICmWD5pcXv1cR6sgOaunIFH9ZSSqNYVNdWMx/yIh8GHclOjfGkc9CGEFolSMny+5OhH0iWZ0t2awynBHKpWBwQB8s+Ln+KAV2Q5UqA0TVDDHPmmwc7G6JrhWwaKqQYtiXpjC86BR0tc7enl5Y2pib6a/sTkXbQ6a5QEKNZgsMlnow2xHbgbdmFe0VVYgYya1XW/NTzCFc4LU0WGiv9rAJpNKAZwxDkgzeqUywqxnq4xN79FIZhh2RBVQ7g99gnRQC+CGFcxF8stKFztt5cwUt7IfmMajW/aP7ARJGjg+IYs0pHwBRcgxHQy1iGzbKDr79EYhQ9IwTOOXwM5RUmjQaNaROYWuqijsQz3EBfbEHown2z6ljoghe3BAldYfNTwrU0V1FVbTeaq0VlwThFZmAmVv4gqCZXaXIBNBE7DSEbFq/gNrBFfgMqZN2WWgVOi0KPKS8umfUTzospKoypEdzpTJoPXSriml03qKqR5HN9FzToHHWlE4/+UGfxyhcHE6FjXUYzhvfspEh9QaKpB5qbb5LhSfBa3max3UK609HKScLUB9FYNvaoX2JARh5LbdzdmZ7EI6ycC4aZEGSSfMUevBgZVQ++1xrhpU1fzBzU6ll2rr+iWhlcnv8mHjSMsJ0KOwgkhUJGA9PfvPNbAFwG7xTiNSnlc2DiqdmzcLO57TjccvL3EIpBLyFvfuPL1WfOglFv03QikSmsYhJTa1xdvisPF8xtx+TkaahIlvG2ws10jb2JJvOaj4A2EaUmidPALW6Y31PRQS5zYv4bBBWvtsC9lMRfizEnsscckHJHBm4gVWlEuNhhJXdtthikgsKJw7G0unoMPO5kDd2tuf/MWTrLHmqeGFsM1elzUSDaZd3OjElscVX+qCHMikdEfpHFnwQntgdyYNelCcVbux1gO3Jx0IduTBdK1Ya4Edc6XEuJXl4QiAxq/H/Yyf2ehdy5R2xnypg4svM3bhxkfg4RG6CGrmZDDOI5BSRJaYHDpzgVOCQM6+sDyaqUUxWN1uzI9mynbZVDlVC5dV/DjVn04ZBfYHDi67oaq4hEZdS5b8tOI8K8/DHPZz1YE8vcrJr+emyinv9GOgvuAap2EGz4JaQsljT6ngJJy29JJgT12SxGMK5PzTlPmhnXyUVhyPx+N4wR2AjDaXy1mH32gmBj1NfUmueR45HLPigr4cj95SpXJY4rLLJ3V+mWoWVeblRydzTqsUvHtsCWwC7wws61Zw/DlH0EN2lzvNAjkxGD5+NuJVnoKpix3ZFcXdfLlpiKBmeZHtJkop0wtQGsQgsaLIltnRgEd8QElrvG5O4PGK0pHv/ax+MzJYcIVDX6+vDzZbeu1tzMvYN3BIfMXqNIPr3RQzbX+lrt9H0j3cAvHrIutu0xzHTiKYXhFymci0poKCPtV8DADEg34ZUm9bA/lBBFc6w10fzpBispbymkLmIch5ZaNfHu8raFeVvgnC64BOTWrfDlPIeEcK19p8/zDosuHtija8dfzgIPke4tAyoGHKp5J2abn368OyljtHYTwYBUl4LQ6bVDPbOnKNTRLUs/l1DDwiD3C+lvnOjPM0GTli0pB5NK4m4gGTwZRezWoWaRwL9uQXFzPeU7u6J39V+5GXmkY5+V6Q76IYgVIP+WcAfs3we013U2+z/yuhg3wujvbxgHt94pdffvnll19++eWXX3755Zdffvnll19++eWXX35ddZlT0uEHfNUH8DD+Y5vWk+AjjP5FS+R9xe+pawWzoVkz8rOZoaCPLMlu67/hkMADeIHQea0J+WYWcYMO5HlNHloJKiZRPDdzI/vjtfHVp/V6fd79P0E0/PZtGEHbsUvPP9nVbx9uSdsN3LY8H07Kenh+bt1Ur7frNmnE//3xoxegP5GIcQqj2zf2538N7HwLbUKUN1Ae5SY+GO3IG866/3u+W6F1ugqRp340LL6fYQ2NF2CQytOXmt6zEjvQ/XKmC9p6oofERIbht2Ejqvj72WzWyEyyYfzhCOF5XfgP96vRjMzOMih+3AJuWHb3zZMuVnj7Ed90jbcXzdj3N8x+eHgAQYKWfU84iQBuj9hNT/1Ane9RQNBX4BxMyBy5Gy4J/mt4ZPb6t+ENQfIPczcj4UsohckPJ8WFpLk5A/UnMmNszj5+0zBVEdhj4e0P2R9NYgxAPt+q4Q1D/vFv3xgFnhsNC/EIrglaD/Aff4nMkMy3fYCdbz/B6a6Xyx/kRmr8e/dO58x2w3fOt6IBdGkywOtBf/etJDl/+fn0m2cfjGbI6g2m3A/QRsuwxSoQt+eW8DA0Ns/azeaKRDcuu2ZN+Q2u8DqEvPXDmHsSnO5mhDXy0rf/ItzqN+CyxW5JJ9CHN2yzrxwkW+HlIUJWO5ULftwKxYSmnNsGfebffcN8lCbQn+EeP/4+xGs8NXD6t+EMN4+TT3FagRBUDWoGOH1nksqbLwD6E+5iFAM3Dc9cAb1sBr1BJT4gxRgeqPFPtnN+4l2xv++RTZ33z/fskLHwTQPoTM6xh3+Pov9QlK0bNzXAO8fUeY17FP1s99zM3IyD+4HgndyTNek/6EyjD4VYgyZzSYIO7Zw+bDByVq9S/g5PDwVwFvp2/3t2z+/NpRzecEPck7fbVRPo/BpD7sEM4a0J9BlBicS+5KHhCr+bPTZGQxQFuDVaAiF9Ap0/EVmvX5Fbm0Enx8T7EI+Na3j1RijN4amRtVeuerQ6XTWDzt85bJb/aFMMGaevmkBnds8R0Nscfa7OuSZqC3f0R7y/4v3+B1JMujUW4Y5ydAR08tjkzkQRuUd3e3Yaq2/BmFhNYmhNLO3UYOmhNmdGxXvDNH0BejOnkzfkYRR19/e/3VT8CTO+rwisyFcAXaD9ziF1Gq+CmdgjM6V4OujCo/pP2TpPp0l3vKkZLrSeZy4m+oFYMSlw02jIrSVk/+HrbuF7DHROoojvq8hJxWfpAcxI30BvCM2stEvF3CJXcGX9MOMBvW/fzuN0Hke9H3472YddcXtZe1X37jivdqcaQVe3zxBbNclmRqNG0IWjzwlw46YiiBjSsEeJIQ06NPDTf4JmDV4R2QhTjD31sGgOzjSZaWvuTn1r5JIWH51HTWC9Wj2LuId7sjJHffhfE+joZYuo3o1T4YqveSJZi05Wrj7z9G/dLgK3Pm56Y8ZxrXzDnOhm4kbv7PVZc+AZYIavz25mZUMgd3bTEm/jL9/MohOt91vgt43H+PAo6U2DAplFEXvtvRF0xGzGrwQNR/2wV1e4Nxo8Mp5LYRcgsxZLtG3ffxavE9KCiQhPNgevV1L7nx0TtP86kdl1IsPt5aMUGzKTBIXRfVM0TH6wKeOjs0nNu17G1NfudzCLE+775KRLG329biX522urZIKf5PWB/HxrfKint1bB9vP19RyC2Gd1rd2YCX9KquwfjXbi69vrujnt+fSG1IHbVn8LH8FNw0esWeDhqWfyddaftj8mpE3SHGH1Kj8e7gulb3nK42MPAme9xDwIbuj9149cyz+weN539vv3b6aOP0+44h3MfE3XZ6gwT4J/h9cBfnIb6zPvgd3B6tFj4Zdffh1f/wMKTPE4TQlnKQAAAABJRU5ErkJggg==";
import {
  Plane, Search, Trash2, Pencil, X, Check, TrendingUp, Ticket, Wallet,
  Calendar, Download, Upload, Building2, Factory, Lock, LogOut, UserPlus, Users, Eye, EyeOff,
  ShieldCheck, Wifi, User, Cloud, Globe2, List, Car, FileText, ArrowLeft,
  MapPin, Compass, Luggage, Anchor, Sparkles, Plus, Printer, SlidersHorizontal, ChevronDown,
  History, Bell, Send, Landmark, Receipt, PieChart, ArrowUpCircle, ArrowDownCircle,
  Banknote, HandCoins, ClipboardList, Globe,
} from "lucide-react";

// A small passport-shaped icon (booklet with a globe emblem) for the Visa section, drawn
// by hand since lucide-react has no built-in "passport" glyph. Mirrors the sizing/stroke
// conventions of the lucide icons it sits alongside (accepts size + className props).
const PassportIcon = ({ size = 22, className = "" }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth={2}
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
  >
    <rect x="5" y="2" width="14" height="20" rx="2" />
    <circle cx="12" cy="10" r="3.2" />
    <path d="M12 6.8v6.4M8.8 10h6.4" />
    <path d="M9 17.5h6" />
  </svg>
);

const MONTHS = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

// ---------- Password hashing ----------
// Employee passwords are stored as a SHA-256 hash (prefixed "sha256:") rather than
// plain text, so anyone who ends up reading the raw stored employee data can't see
// the actual password. Hashing happens client-side with the browser's built-in
// Web Crypto API — no server involved.
const HASH_PREFIX = "sha256:";

const hashPassword = async (plain) => {
  const data = new TextEncoder().encode(plain);
  const digest = await crypto.subtle.digest("SHA-256", data);
  const hex = Array.from(new Uint8Array(digest)).map((b) => b.toString(16).padStart(2, "0")).join("");
  return HASH_PREFIX + hex;
};

// Compares a plain-text password the user just typed against a stored value. Supports
// both the new hashed format and legacy plain-text entries (e.g. from an older backup),
// so nothing breaks for accounts created before this change.
const verifyPassword = async (storedValue, plainAttempt) => {
  if (typeof storedValue !== "string") return false;
  if (storedValue.startsWith(HASH_PREFIX)) {
    return (await hashPassword(plainAttempt)) === storedValue;
  }
  return storedValue === plainAttempt; // legacy plain-text account
};

const monthKey = (dateStr) => (dateStr ? dateStr.slice(0, 7) : "No date");
// Storage stays in the native YYYY-MM-DD format (required by <input type="date">),
// but everywhere we display the date to the user we show it as DD-MMM-YYYY, with the
// month written as its first three letters, capitalized (e.g. "03-AUG-2026").
const formatDisplayDate = (dateStr) => {
  if (!dateStr) return "";
  const [y, m, d] = dateStr.split("-");
  if (!y || !m || !d) return dateStr;
  const monthAbbr = (MONTHS[parseInt(m, 10) - 1] || m).slice(0, 3).toUpperCase();
  return `${d}-${monthAbbr}-${y}`;
};
const monthLabel = (key) => {
  if (key === "No date") return key;
  const [y, m] = key.split("-");
  const idx = parseInt(m, 10) - 1;
  return `${MONTHS[idx] || m} ${y}`;
};

// Formats an ISO timestamp as DD-MMM-YYYY HH:MM for showing when a note edit happened.
const formatDateTime = (iso) => {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return "";
  const dd = String(d.getDate()).padStart(2, "0");
  const monthAbbr = MONTHS[d.getMonth()].slice(0, 3).toUpperCase();
  const yyyy = d.getFullYear();
  const hh = String(d.getHours()).padStart(2, "0");
  const min = String(d.getMinutes()).padStart(2, "0");
  return `${dd}-${monthAbbr}-${yyyy} ${hh}:${min}`;
};

// ==================== License / Activation ====================
// This app requires an activation code before it can be used at all.
// Add, remove, or edit codes below yourself — no coding needed beyond this list.
//   - expiresAt: null      -> the code works forever (permanent license)
//   - expiresAt: "YYYY-MM-DD" -> the code stops working after that date (subscription-style)
// You can list as many codes as you like — e.g. a different one per client, or one
// permanent code plus a few time-limited trial codes.
const LICENSE_KEYS = [
  { code: "PERLA-DIMARE-2026", expiresAt: null },
  { code: "PERLA-TRIAL-30D", expiresAt: "2026-09-04" },
];

// The key under which the activated code is remembered in this browser, so the
// activation screen doesn't reappear every time the app is opened.
const LICENSE_STORAGE_KEY = "ftm_license_activation";

const checkLicenseCode = (rawCode) => {
  const code = (rawCode || "").trim().toUpperCase();
  if (!code) return { valid: false, reason: "Please enter an activation code" };
  const entry = LICENSE_KEYS.find((l) => l.code.trim().toUpperCase() === code);
  if (!entry) return { valid: false, reason: "Invalid activation code" };
  if (entry.expiresAt) {
    const todayStr = new Date().toISOString().slice(0, 10);
    if (todayStr > entry.expiresAt) {
      return { valid: false, reason: `This activation code expired on ${entry.expiresAt}` };
    }
  }
  return { valid: true, code: entry.code, expiresAt: entry.expiresAt };
};

const emptyCustomerRow = () => ({ name: "", ticketNumber: "", conjunction: false, ticketNumber2: "", pnrReference: "" });

// Ticket supplier / booking source options.
const SUPPLIERS = ["Amadeus", "Sabre", "NDC", "Lowcost"];

const CAR_TYPES = ["Sedan", "Mini Van", "H1", "Coaster", "Bus"];

// Hour/minute option lists for the smooth, pill-shaped transfer time picker below
// (two plain <select>s fused into one field instead of the clunky native <input type="time">).
const TIME_HOURS = Array.from({ length: 24 }, (_, i) => String(i).padStart(2, "0"));
const TIME_MINUTES = Array.from({ length: 60 }, (_, i) => String(i).padStart(2, "0"));

// A single "HH:MM" value rendered as two borderless selects (hour, minute) inside one
// shared pill, so it reads as one smooth control instead of two separate boxes — used
// for transfer pickup times in the Transportation section.
const TimeSelect = ({ value, onChange }) => {
  const [h = "", m = ""] = (value || "").split(":");
  const update = (nh, nm) => {
    if (!nh && !nm) { onChange(""); return; }
    onChange(`${nh || "00"}:${nm || "00"}`);
  };
  return (
    <div className="w-full flex items-center border border-stone-300 rounded-xl px-3 py-2 text-sm focus-within:outline-none focus-within:ring-2 focus-within:ring-teal-700 bg-white">
      <select
        aria-label="Hour"
        className="flex-1 bg-transparent focus:outline-none appearance-none text-center"
        value={h}
        onChange={(e) => update(e.target.value, m)}
      >
        <option value="">--</option>
        {TIME_HOURS.map((v) => (
          <option key={v} value={v}>{v}</option>
        ))}
      </select>
      <span className="text-stone-400 px-0.5">:</span>
      <select
        aria-label="Minute"
        className="flex-1 bg-transparent focus:outline-none appearance-none text-center"
        value={m}
        onChange={(e) => update(h, e.target.value)}
      >
        <option value="">--</option>
        {TIME_MINUTES.map((v) => (
          <option key={v} value={v}>{v}</option>
        ))}
      </select>
    </div>
  );
};

// Saved companies were originally plain strings; this reads the name whether an entry
// is still a legacy string or the newer { name, taxNumber, commercialReg, phones } record.
const companyName = (c) => (typeof c === "string" ? c : (c && c.name) || "");

const emptyCompanyDraft = { name: "", taxNumber: "", commercialReg: "", phones: "" };

// Local YYYY-MM-DD for today, matching the native <input type="date"> format.
const todayDateStr = () => {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
};

// ============================================================
// ---------- Accounts (accounting module) ----------
// ============================================================
// Expense categories offered on the expense form — covers the recurring cost lines a
// travel agency typically tracks. "Other" always stays free-text via a separate field.
const EXPENSE_CATEGORIES = [
  "إيجار",
  "مرتبات وعمولات",
  "كهرباء ومياه",
  "اتصالات وإنترنت",
  "تسويق وإعلان",
  "صيانة",
  "ضرائب ورسوم حكومية",
  "بنك وتحويلات",
  "قرطاسية ومطبوعات",
  "أخرى",
];
const TREASURY_ACCOUNT_TYPES = [
  { value: "cash", label: "خزينة نقدية" },
  { value: "bank", label: "حساب بنكي" },
];
// Categories for manual treasury entries (money moved without a linked customer/
// supplier/expense record — capital injections, owner drawings, transfers, etc.).
const TREASURY_ENTRY_CATEGORIES_IN = ["رأس مال / إيداع مالك", "تحويل من حساب آخر", "إيرادات أخرى"];
const TREASURY_ENTRY_CATEGORIES_OUT = ["مسحوبات مالك", "تحويل إلى حساب آخر", "مصروفات أخرى"];

// ---------- Accounts section English translations ----------
// The Accounts module supports an Arabic/English display toggle. Underlying stored
// values (expense categories, treasury entry categories, etc.) stay in Arabic for data
// consistency; these maps only translate what's shown on screen.
const EXPENSE_CATEGORY_LABELS_EN = {
  "إيجار": "Rent",
  "مرتبات وعمولات": "Salaries & Commissions",
  "كهرباء ومياه": "Electricity & Water",
  "اتصالات وإنترنت": "Communications & Internet",
  "تسويق وإعلان": "Marketing & Advertising",
  "صيانة": "Maintenance",
  "ضرائب ورسوم حكومية": "Taxes & Government Fees",
  "بنك وتحويلات": "Bank & Transfers",
  "قرطاسية ومطبوعات": "Stationery & Printing",
  "أخرى": "Other",
};
const TREASURY_ACCOUNT_TYPE_LABELS_EN = { cash: "Cash Treasury", bank: "Bank Account" };
const TREASURY_ENTRY_CATEGORY_LABELS_EN = {
  "رأس مال / إيداع مالك": "Capital / Owner Deposit",
  "تحويل من حساب آخر": "Transfer From Another Account",
  "إيرادات أخرى": "Other Income",
  "مسحوبات مالك": "Owner Withdrawals",
  "تحويل إلى حساب آخر": "Transfer To Another Account",
  "مصروفات أخرى": "Other Expenses",
};
const ACCOUNTS_I18N = {
  ar: {
    tabOverview: "نظرة عامة", tabSuppliers: "الموردين", tabCustomers: "العملاء",
    tabTreasury: "الخزينة والبنوك", tabExpenses: "المصروفات", tabReports: "التقارير المالية",
    monthRevenue: "أرباح الشهر الحالى", monthExpenses: "مصروفات الشهر الحالى",
    monthNetProfit: "صافى ربح الشهر", totalTreasuryBalance: "إجمالى رصيد الخزينة والبنوك",
    totalOwedSuppliers: "مستحق للموردين", totalOwedCustomers: "مستحق من العملاء",
    profitBySection: "أرباح الشهر حسب القسم",
    searchSupplier: "ابحث عن مورد...", searchCustomer: "ابحث عن عميل...",
    colSupplier: "المورد", colCustomer: "العميل", colSections: "الأقسام",
    colTotalOwed: "إجمالى المستحق", colTotalDue: "إجمالى المستحق",
    colPaid: "المدفوع", colCollected: "المحصل", colRemaining: "المتبقى",
    noSuppliers: "لا يوجد موردون", noCustomers: "لا يوجد عملاء",
    accountsAndTreasuries: "الحسابات والخزائن", addAccount: "إضافة حساب",
    noAccountsYet: "لا توجد حسابات بعد. أضف خزينة نقدية أو حساب بنكى للبدء.",
    treasuryMovement: "حركة الخزينة", allAccounts: "كل الحسابات", manualEntry: "قيد يدوى",
    colDate: "التاريخ", colAccount: "الحساب", colStatement: "البيان", colAmount: "المبلغ",
    noTransactions: "لا توجد حركات", allCategories: "كل التصنيفات", addExpense: "إضافة مصروف",
    colCategory: "التصنيف", colDescription: "الوصف", noExpenses: "لا توجد مصروفات",
    rangeToday: "اليوم", rangeMonth: "الشهر الحالى", rangeCustom: "مخصص", to: "إلى",
    exportExcel: "تصدير Excel", revenueOf: (s) => `إيرادات ${s}`, bookingsCount: "حجز",
    totalRevenue: "إجمالى الإيرادات", totalExpenses: "إجمالى المصروفات", netProfit: "صافى الربح",
    expensesByCategory: "المصروفات حسب التصنيف", noExpensesInPeriod: "لا توجد مصروفات فى هذه الفترة",
    editExpense: "تعديل مصروف", descriptionOptional: "الوصف (اختيارى)", amountEgp: "المبلغ (ج.م)",
    payFromAccount: "صرف من (خزينة/حساب)", selectAccount: "اختر حساب", notes: "ملاحظات",
    saveChanges: "حفظ التعديل",
    editAccount: "تعديل حساب", addAccountTreasury: "إضافة حساب/خزينة", accountName: "اسم الحساب",
    accountNamePlaceholder: "مثال: خزينة المكتب، حساب بنك مصر", type: "النوع",
    openingBalance: "الرصيد الافتتاحى (ج.م)",
    manualEntryTitle: "قيد يدوى", directionIn: "وارد (+)", directionOut: "منصرف (-)",
    item: "البند", saveEntry: "حفظ القيد",
    recordNewPayment: "تسجيل دفعة جديدة", payFrom: "ادفع من (خزينة/حساب)",
    notesOptional: "ملاحظات (اختيارى)", recordPayment: "تسجيل الدفعة",
    paymentHistory: "سجل المدفوعات", noPaymentsRecorded: "لا توجد مدفوعات مسجلة",
    relatedBookings: "الحجوزات المرتبطة", noBookings: "لا توجد حجوزات",
    recordNewCollection: "تسجيل تحصيل جديد", collectInto: "التحصيل فى (خزينة/حساب)",
    recordCollection: "تسجيل التحصيل", collectionHistory: "سجل التحصيلات",
    noCollectionsRecorded: "لا توجد تحصيلات مسجلة",
    currency: "ج.م",
  },
  en: {
    tabOverview: "Overview", tabSuppliers: "Suppliers", tabCustomers: "Customers",
    tabTreasury: "Treasury & Banks", tabExpenses: "Expenses", tabReports: "Financial Reports",
    monthRevenue: "This Month's Profit", monthExpenses: "This Month's Expenses",
    monthNetProfit: "Net Profit This Month", totalTreasuryBalance: "Total Treasury & Bank Balance",
    totalOwedSuppliers: "Owed to Suppliers", totalOwedCustomers: "Owed by Customers",
    profitBySection: "This Month's Profit by Section",
    searchSupplier: "Search suppliers...", searchCustomer: "Search customers...",
    colSupplier: "Supplier", colCustomer: "Customer", colSections: "Sections",
    colTotalOwed: "Total Owed", colTotalDue: "Total Due",
    colPaid: "Paid", colCollected: "Collected", colRemaining: "Remaining",
    noSuppliers: "No suppliers found", noCustomers: "No customers found",
    accountsAndTreasuries: "Accounts & Treasuries", addAccount: "Add Account",
    noAccountsYet: "No accounts yet. Add a cash treasury or bank account to get started.",
    treasuryMovement: "Treasury Transactions", allAccounts: "All Accounts", manualEntry: "Manual Entry",
    colDate: "Date", colAccount: "Account", colStatement: "Description", colAmount: "Amount",
    noTransactions: "No transactions found", allCategories: "All Categories", addExpense: "Add Expense",
    colCategory: "Category", colDescription: "Description", noExpenses: "No expenses found",
    rangeToday: "Today", rangeMonth: "This Month", rangeCustom: "Custom", to: "to",
    exportExcel: "Export Excel", revenueOf: (s) => `${s} Revenue`, bookingsCount: "bookings",
    totalRevenue: "Total Revenue", totalExpenses: "Total Expenses", netProfit: "Net Profit",
    expensesByCategory: "Expenses by Category", noExpensesInPeriod: "No expenses in this period",
    editExpense: "Edit Expense", descriptionOptional: "Description (optional)", amountEgp: "Amount (EGP)",
    payFromAccount: "Paid From (Treasury/Account)", selectAccount: "Select account", notes: "Notes",
    saveChanges: "Save Changes",
    editAccount: "Edit Account", addAccountTreasury: "Add Account/Treasury", accountName: "Account Name",
    accountNamePlaceholder: "e.g. Office Treasury, Bank Misr Account", type: "Type",
    openingBalance: "Opening Balance (EGP)",
    manualEntryTitle: "Manual Entry", directionIn: "In (+)", directionOut: "Out (-)",
    item: "Item", saveEntry: "Save Entry",
    recordNewPayment: "Record New Payment", payFrom: "Pay From (Treasury/Account)",
    notesOptional: "Notes (optional)", recordPayment: "Record Payment",
    paymentHistory: "Payment History", noPaymentsRecorded: "No payments recorded",
    relatedBookings: "Related Bookings", noBookings: "No bookings",
    recordNewCollection: "Record New Collection", collectInto: "Collect Into (Treasury/Account)",
    recordCollection: "Record Collection", collectionHistory: "Collection History",
    noCollectionsRecorded: "No collections recorded",
    currency: "EGP",
  },
};

const getEmptyExpenseForm = () => ({
  id: null,
  date: todayDateStr(),
  category: EXPENSE_CATEGORIES[0],
  description: "",
  amount: "",
  accountId: "",
  note: "",
});
const getEmptyTreasuryAccountForm = () => ({ id: null, name: "", type: "cash", openingBalance: "" });
const getEmptySupplierPaymentForm = () => ({
  id: null,
  date: todayDateStr(),
  supplier: "",
  amount: "",
  accountId: "",
  method: "cash",
  note: "",
});
const getEmptyCustomerPaymentForm = () => ({
  id: null,
  date: todayDateStr(),
  customer: "",
  amount: "",
  accountId: "",
  method: "cash",
  note: "",
});
const getEmptyTreasuryEntryForm = () => ({
  id: null,
  date: todayDateStr(),
  accountId: "",
  direction: "in",
  category: TREASURY_ENTRY_CATEGORIES_IN[0],
  amount: "",
  note: "",
});

// A function (not a static object) so every new/reset ticket picks up TODAY'S date
// at the moment it's created, rather than whatever date happened to be "today" when
// the app first loaded. The user can still change it manually afterward.
const getEmptyForm = () => ({
  id: null,
  employee: "",
  company: "",
  supplier: "",
  customersCount: 1,
  customers: [emptyCustomerRow()],
  from: "",
  to: "",
  // Return-leg airport, shown only for a round trip — the outbound "to" airport is
  // usually where the return departs from, but this lets it be entered separately
  // when it differs (e.g. a different city on the way back).
  returnAirport: "",
  // Multi-destination (multi-city) route support: when multiDestination is on, the
  // route is described as an ordered list of stops (e.g. ["CAI","DXB","BKK"]) instead
  // of a single from/to pair. "from"/"to" are still kept in sync (first/last stop) so
  // every place that reads a plain origin/destination keeps working unchanged.
  multiDestination: false,
  destinations: ["", ""],
  // Trip type shown next to the multi-destination toggle: "oneWay" or "roundTrip".
  tripType: "oneWay",
  airline: "",
  date: todayDateStr(),
  netPrice: "",
  soldPrice: "",
  notes: "",
  // Reissue tracking: when isReissued is on, oldTicketNumber is looked up against
  // existing tickets to auto-fill oldTicketIssueDate and every other field below
  // (company, supplier, route, airline, prices, customer names) from that old ticket.
  isReissued: false,
  oldTicketNumber: "",
  oldTicketIssueDate: "",
  // Refund tracking: a list of refund records (each with two amounts — refunded by the
  // airline, refunded to the customer), entered right in the ticket form next to the
  // reissue box. Empty while nothing's been refunded. A single booking can have several
  // customers/tickets, so this is a list — one entry per refunded customerIndex — rather
  // than a single object, so refunding more than one ticket on the same booking doesn't
  // overwrite an earlier one.
  refunds: [],
});

// Renders a ticket's route as a single "A → B" (or "A → B → C → ..." for a
// multi-destination/multi-city booking) string for lists, detail views, and exports.
const routeLabel = (t) => {
  const stops = Array.isArray(t.destinations) ? t.destinations.map((d) => (d || "").trim()).filter(Boolean) : [];
  if (t.multiDestination && stops.length >= 2) return stops.join(" → ");
  return `${t.from || "-"} → ${t.to || "-"}`;
};

// Room types offered on a hotel booking's room line.
const ROOM_TYPES = [
  { value: "single", label: "Single" },
  { value: "double", label: "Double" },
  { value: "triple", label: "Triple" },
];

// Meal plan offered on a hotel booking's room line.
const MEAL_PLANS = [
  { value: "ro", label: "Room Only" },
  { value: "bb", label: "Bed & Breakfast" },
  { value: "hb", label: "Half Board" },
  { value: "fb", label: "Full Board" },
  { value: "ai", label: "All Inclusive" },
];

// Max number of adult guests a room type can hold — drives how many guest-name fields
// are shown for a room line (Single -> 1, Double -> 2, Triple -> 3).
const ROOM_CAPACITY = { single: 1, double: 2, triple: 3 };

// A single adult guest staying in a room.
const emptyGuest = () => ({
  id: `G-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
  name: "",
});

// A child staying in a room, with an age (in whole years, 0–11) alongside the name.
const emptyChild = () => ({
  id: `C-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
  name: "",
  age: "",
});

// Converts Arabic-Indic (٠-٩) and Extended Arabic-Indic (۰-۹) digits to standard 0-9,
// then strips anything that isn't a digit. Using type="text" with this instead of
// type="number" avoids the age field silently rejecting keystrokes on Arabic keyboards,
// which type="number" does with non-Latin digits.
const sanitizeAgeInput = (raw) => {
  let v = raw
    .replace(/[\u0660-\u0669]/g, (d) => String(d.charCodeAt(0) - 0x0660))
    .replace(/[\u06F0-\u06F9]/g, (d) => String(d.charCodeAt(0) - 0x06F0));
  v = v.replace(/[^0-9]/g, "");
  if (v !== "" && parseInt(v, 10) > 11) v = "11";
  return v;
};

// Resizes a room line's guest list to match its room type's capacity, keeping any
// names already entered and padding/truncating as needed.
const guestsForCapacity = (guests, capacity) => {
  const list = (Array.isArray(guests) ? guests : []).slice(0, capacity).map((g) => ({ ...g }));
  while (list.length < capacity) list.push(emptyGuest());
  return list;
};

const HOTEL_CURRENCIES = [
  { value: "EGP", label: "EGP" },
  { value: "USD", label: "USD" },
];

// A single room line within a hotel booking: a room type + meal plan combination, its own
// currency, count, and net/sold price per room per night — e.g. "1x Single, Half Board,
// EGP" and "2x Double, All Inclusive, USD" can both live inside the same booking.
const emptyRoomLine = () => ({
  id: `RL-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
  roomType: "single",
  mealPlan: "bb",
  currency: "EGP",
  count: 1,
  netPrice: "",
  soldPrice: "",
  // Each room now carries its own stay dates, since different rooms on the same
  // booking can check in/out on different days.
  checkIn: todayDateStr(),
  checkOut: todayDateStr(),
  // Adult guest names — sized to the default room type's capacity (single -> 1).
  guests: guestsForCapacity([], ROOM_CAPACITY.single),
  // Children staying in this room, each with a name and age (0–11 years).
  children: [],
});

// A function (not a static object) so every new/reset hotel booking picks up TODAY'S
// date at the moment it's created, same rationale as getEmptyForm() above.
const getEmptyHotelForm = () => ({
  id: null,
  employee: "",
  customer: "",
  hotel: "",
  supplier: "",
  roomLines: [emptyRoomLine()],
  // The date the reservation itself was made — separate from each room's own
  // check-in/check-out dates below.
  bookingDate: todayDateStr(),
  notes: "",
});

// A single customer on a visa booking — just a name (visas don't track per-customer
// ticket numbers the way flight tickets do).
const emptyVisaCustomer = () => ({ name: "" });

// Fills/trims a visa booking's customer list to match the requested count, keeping any
// names already entered — same rationale as resizeCustomers() above.
const resizeVisaCustomers = (customers, count) => {
  const n = Math.max(1, Math.min(50, parseInt(count, 10) || 1));
  const next = [...customers];
  while (next.length < n) next.push(emptyVisaCustomer());
  next.length = n;
  return next;
};

// A function (not a static object) so every new/reset visa booking picks up TODAY'S
// date at the moment it's created, same rationale as getEmptyHotelForm() above.
const getEmptyVisaForm = () => ({
  id: null,
  customersCount: 1,
  customers: [emptyVisaCustomer()],
  visaType: "",
  supplier: "",
  currency: "EGP",
  netPrice: "",
  soldPrice: "",
  bookingDate: todayDateStr(),
});

// A function (not a static object) so every new/reset transfer booking picks up TODAY'S
// date at the moment it's created, same rationale as getEmptyVisaForm() above.
const getEmptyCarForm = () => ({
  id: null,
  customerName: "",
  phone: "",
  routeFrom: "",
  routeTo: "",
  carType: "",
  supplier: "",
  hasWaiting: false,
  waitingHours: "",
  isRoundTrip: false,
  driverTip: "",
  startsAtAirport: false,
  flightNumber: "",
  currency: "EGP",
  netPrice: "",
  soldPrice: "",
  bookingDate: todayDateStr(),
  bookingTime: "",
  returnDate: "",
  returnTime: "",
  entryDate: todayDateStr(),
  collection: "",
});

// Given a ticket number like "077-1234567890", returns the same prefix with the numeric
// part increased by one, keeping the same digit width (e.g. "077-1234567891").
// Returns "" if the ticket number doesn't match the expected PREFIX-DIGITS shape.
// Auto-sequencing only ever advances the LAST THREE digits of the serial number (wrapping
// 999 back to 000); everything before them — including the rest of the serial — stays fixed,
// since that part identifies the batch/booking rather than the individual ticket.
const nextTicketNumber = (ticketNumber) => {
  if (!ticketNumber) return "";
  const match = ticketNumber.match(/^([A-Z0-9]{3})-(\d+)$/);
  if (!match) return "";
  const [, prefix, digits] = match;
  if (digits.length <= 3) {
    const wrapped = ((parseInt(digits, 10) + 1) % (10 ** digits.length)).toString().padStart(digits.length, "0");
    return `${prefix}-${wrapped}`;
  }
  const head = digits.slice(0, -3);
  const tail = digits.slice(-3);
  const nextTail = ((parseInt(tail, 10) + 1) % 1000).toString().padStart(3, "0");
  return `${prefix}-${head}${nextTail}`;
};

// Given a ticket number's last three digits, returns the "-XXX" suffix used for a
// conjunction ticket — the customer's second ticket number issued together with the
// first, which airlines write as just the incremented tail after a dash (e.g. ticket
// "077-1234567890" gets a conjunction suffix of "-891"). Wraps 999 back to 000, same as
// nextTicketNumber above. Returns "" if there aren't at least three digits to work from.
const conjunctionTicketSuffix = (ticketNumber) => {
  const digits = (ticketNumber || "").replace(/[^0-9]/g, "");
  if (digits.length < 3) return "";
  const tail = digits.slice(-3);
  const nextTail = ((parseInt(tail, 10) + 1) % 1000).toString().padStart(3, "0");
  return `-${nextTail}`;
};

// Given a customer row, returns the ticket number the NEXT customer's auto-sequenced
// number should be generated from. If this customer has a conjunction (second) ticket,
// that second number was already issued to them, so the next customer continues after
// its tail rather than after the first ticket's tail — e.g. first ticket
// "077-1234567890" with a conjunction suffix of "-891" means the next customer should
// get "077-1234567892", not "077-1234567891" (which is this customer's own conjunction
// ticket). Falls back to the plain ticket number when there's no conjunction ticket.
const lastIssuedTicketNumber = (customer) => {
  if (!customer) return "";
  if (customer.conjunction && customer.ticketNumber2) {
    const match = (customer.ticketNumber || "").match(/^([A-Z0-9]{3})-(\d+)$/);
    const tailDigits = customer.ticketNumber2.replace(/[^0-9]/g, "");
    if (match && tailDigits) {
      const [, prefix, num] = match;
      const head = num.length > 3 ? num.slice(0, -3) : "";
      return `${prefix}-${head}${tailDigits.padStart(3, "0")}`;
    }
  }
  return customer.ticketNumber;
};

// Fills/trims the customers array to match the requested count, keeping existing entries
const resizeCustomers = (customers, count) => {
  const n = Math.max(1, Math.min(50, parseInt(count, 10) || 1));
  const next = [...customers];
  while (next.length < n) next.push(emptyCustomerRow());
  next.length = n;
  return next;
};

// Job grades shown to the main account when creating/editing an employee. Picking a
// grade fills in a sensible starting set of permission toggles below (see
// ROLE_PRESETS), but every toggle can still be switched on or off by hand afterwards —
// the grade is a starting point/label, not a lock. Grade is purely descriptive; access
// is always driven by the individual toggles stored on the employee record.
//
// Alongside the four original, all-section grades (Supervisor/Employee/
// Accountant/Owner — the all-section Manager grade has been removed), there's a grade
// per section per tier — e.g. "Flights Employee", "Tourism Supervisor", "Transportation
// Manager" — that automatically limits the employee to that one section (every other
// section is switched off) with a sensible starting permission level for that tier.
// Main account/Owner can still fine-tune or widen access afterward from the
// Permissions screen; picking the grade just sets the starting point.
const SECTION_ROLE_LABELS = { flights: "Flights", hotels: "Hotels", visa: "Visa", cars: "Transportation", files: "Files" };
// Custom, job-title-style labels for the per-section Employee grade specifically.
const EMPLOYEE_GRADE_LABELS = {
  flights: "Ticketing Agent",
  hotels: "Tourism Employee",
  visa: "Visa Employee",
  cars: "Transportation Employee",
};
// Same idea for the per-section Supervisor grade — "Ticketing Supervisor" instead of
// "Flights Supervisor", "Tourism Supervisor" instead of "Hotels Supervisor".
const SUPERVISOR_GRADE_LABELS = {
  flights: "Ticketing Supervisor",
  hotels: "Tourism Supervisor",
  visa: "Visa Supervisor",
  cars: "Transportation Supervisor",
};
// Same idea for the per-section Manager grade — "Ticketing Manager" instead of
// "Flights Manager", "Tourism Manager" instead of "Hotels Manager".
const MANAGER_GRADE_LABELS = {
  flights: "Ticketing Manager",
  hotels: "Tourism Manager",
  visa: "Visa Manager",
  cars: "Transportation Manager",
};
// Every section that has bookings tagged with an owning employee (see
// SECTIONS_WITH_OWNERSHIP) gets its own dedicated Employee, Supervisor, and Manager
// grade — Files is the one exception for all three, since it's normally reached
// through the other sections' "Copy to a file" action rather than worked directly by
// an Employee, Supervisor, or Manager.
const SECTIONS_WITH_EMPLOYEE_GRADE = ["flights", "hotels", "visa", "cars"];
const SECTIONS_WITH_SUPERVISOR_GRADE = ["flights", "hotels", "visa", "cars"];
const SECTIONS_WITH_MANAGER_GRADE = ["flights", "hotels", "visa", "cars"];
const ALL_ROLE_SECTIONS = ["flights", "hotels", "visa", "cars", "files"];
const sectionOnlyAccess = (section) => ({ flights: false, hotels: false, visa: false, cars: false, files: false, [section]: true });
const SECTION_TIER_PERMS = {
  employee: { canViewAll: false, canEdit: false, canDelete: false, canManageCompanies: false },
  supervisor: { canViewAll: true, canEdit: true, canDelete: false, canManageCompanies: false },
  manager: { canViewAll: true, canEdit: true, canDelete: true, canManageCompanies: true },
};
const sectionRolePreset = (section, tier) => {
  const p = SECTION_TIER_PERMS[tier];
  return {
    canViewAll: p.canViewAll,
    canAdd: true,
    canEdit: p.canEdit,
    canDelete: p.canDelete,
    isAccounting: false,
    canManageCompanies: p.canManageCompanies,
    isOwner: false,
    sections: sectionOnlyAccess(section),
    sectionPerms: { [section]: { canViewAll: p.canViewAll, canEdit: p.canEdit, canDelete: p.canDelete } },
  };
};

const EMPLOYEE_ROLES = [
  { value: "supervisor", label: "Supervisor" },
  { value: "employee", label: "Employee" },
  { value: "owner", label: "Owner" },
  { value: "gm", label: "GM" },
  { value: "accountant", label: "Accountant" },
  ...SECTIONS_WITH_EMPLOYEE_GRADE.map((s) => ({ value: `employee_${s}`, label: EMPLOYEE_GRADE_LABELS[s] || `${SECTION_ROLE_LABELS[s]} Employee` })),
  ...SECTIONS_WITH_SUPERVISOR_GRADE.map((s) => ({ value: `supervisor_${s}`, label: SUPERVISOR_GRADE_LABELS[s] || `${SECTION_ROLE_LABELS[s]} Supervisor` })),
  ...SECTIONS_WITH_MANAGER_GRADE.map((s) => ({ value: `manager_${s}`, label: MANAGER_GRADE_LABELS[s] || `${SECTION_ROLE_LABELS[s]} Manager` })),
];

// Starting toggle values applied when a grade is picked. All toggles are then freely
// editable by hand, independent of which grade is selected. "Owner" and "GM" are a
// step above the rest: full ticket/company access like the other grades below, PLUS
// admin-level access to Manage employees and Backup/Restore (granted separately via
// isOwner, checked alongside currentUser.isAdmin wherever those are gated) — the one
// thing an Owner or GM never gets is the License panel, which stays reserved for true
// main accounts. GM is defined as an exact copy of the Owner preset — same starting
// permissions, same isOwner flag — it's simply a second, identically-privileged grade.
// The five original grades are explicitly all-section (sections + sectionPerms reset to
// full access) so switching *back* to one of them from a section-limited grade restores
// every section, rather than leaving the old restriction in place.
const ALL_SECTIONS_ON = { flights: true, hotels: true, visa: true, cars: true, files: true };
const ROLE_PRESETS = {
  supervisor: { canViewAll: true, canAdd: true, canEdit: true, canDelete: false, isAccounting: false, canManageCompanies: false, isOwner: false, sections: { ...ALL_SECTIONS_ON }, sectionPerms: {} },
  employee: { canViewAll: false, canAdd: true, canEdit: false, canDelete: false, isAccounting: false, canManageCompanies: false, isOwner: false, sections: { ...ALL_SECTIONS_ON }, sectionPerms: {} },
  accountant: { canViewAll: true, canAdd: false, canEdit: false, canDelete: false, isAccounting: true, canManageCompanies: false, isOwner: false, sections: { ...ALL_SECTIONS_ON }, sectionPerms: {} },
  owner: { canViewAll: true, canAdd: true, canEdit: true, canDelete: true, isAccounting: false, canManageCompanies: true, isOwner: true, sections: { ...ALL_SECTIONS_ON }, sectionPerms: {} },
  gm: { canViewAll: true, canAdd: true, canEdit: true, canDelete: true, isAccounting: false, canManageCompanies: true, isOwner: true, sections: { ...ALL_SECTIONS_ON }, sectionPerms: {} },
  ...Object.fromEntries(SECTIONS_WITH_EMPLOYEE_GRADE.map((s) => [`employee_${s}`, sectionRolePreset(s, "employee")])),
  ...Object.fromEntries(SECTIONS_WITH_SUPERVISOR_GRADE.map((s) => [`supervisor_${s}`, sectionRolePreset(s, "supervisor")])),
  ...Object.fromEntries(SECTIONS_WITH_MANAGER_GRADE.map((s) => [`manager_${s}`, sectionRolePreset(s, "manager")])),
};

const roleLabel = (value) => (EMPLOYEE_ROLES.find((r) => r.value === value) || {}).label || "Employee";

// Grade picker on the Add employee page groups the 17 grades into three per-tier
// dropdowns (Manager / Supervisor / Employee). Supervisor and Employee each also hold
// a general, all-section grade alongside their per-department variants; Manager no
// longer has a general grade, so its dropdown holds only the four per-department
// variants. Owner, GM, and Accountant stand alone next to the three dropdowns since
// none of them has department-specific variants.
const MANAGER_GRADES = EMPLOYEE_ROLES.filter((r) => r.value === "manager" || r.value.startsWith("manager_"));
const SUPERVISOR_GRADES = EMPLOYEE_ROLES.filter((r) => r.value === "supervisor" || r.value.startsWith("supervisor_"));
const EMPLOYEE_GRADES = EMPLOYEE_ROLES.filter((r) => r.value === "employee" || r.value.startsWith("employee_"));
const GRADE_TIER_GROUPS = [
  { key: "manager", title: "Manager", roles: MANAGER_GRADES },
  { key: "supervisor", title: "Supervisor", roles: SUPERVISOR_GRADES },
  { key: "employee", title: "Employee", roles: EMPLOYEE_GRADES },
];

// Which of the app's sections (Flights/Hotels/Visa/Transportation/Files) an employee can
// see and use, independent of their ticket permissions (view/add/edit/delete) above — an
// employee could, for example, be allowed to add tickets but only in the Hotels section.
// Every existing employee predates this feature, so any section missing from a stored
// record is treated as allowed (see employeeSections below) rather than silently locking
// people out of sections they already had access to.
const SECTION_OPTIONS = [
  { value: "flights", label: "Flights" },
  { value: "hotels", label: "Hotels" },
  { value: "visa", label: "Visa" },
  { value: "cars", label: "Transportation" },
  { value: "files", label: "Files" },
];
const DEFAULT_SECTIONS = { flights: true, hotels: true, visa: true, cars: true, files: true };
// Merges an employee's stored section toggles over the all-allowed defaults, so a
// legacy record with no "sections" field at all (or missing individual keys) still
// resolves to full access rather than blocking every section.
const employeeSections = (emp) => ({ ...DEFAULT_SECTIONS, ...((emp && emp.sections) || {}) });

// Fine-grained permissions — View all services / Edit / Delete — set independently for
// each of the five sections, on top of the on/off "section access" toggle above. All
// three are fully independent switches: an employee can have Edit on with View all
// services off, which simply limits them to editing their own records. Every section —
// Flights, Hotels, Visa, Transportation, and Files — tracks which employee created each
// record, so "View all services" has real meaning everywhere (off = only your own
// records).
const SECTIONS_WITH_OWNERSHIP = ["flights", "hotels", "visa", "cars", "files"];
// Every existing employee predates this feature, so a section with nothing stored in
// sectionPerms falls back to that employee's old, single account-wide
// canViewAll/canEdit/canDelete values — permissions stay exactly as they were until the
// main account deliberately customizes a specific section.
const employeeSectionPerm = (emp, section) => {
  const legacy = {
    canViewAll: !!(emp && (emp.canViewAll || emp.canEdit || emp.canDelete)),
    canEdit: !!(emp && emp.canEdit),
    canDelete: !!(emp && emp.canDelete),
  };
  const stored = (emp && emp.sectionPerms && emp.sectionPerms[section]) || {};
  const merged = { ...legacy, ...stored };
  if (emp && emp.isAccounting) return { canViewAll: true, canEdit: false, canDelete: false };
  return merged;
};

// Applies the coherence rules that keep the six permission toggles consistent with
// each other, no matter which one was just changed by hand:
// - Editing, deleting, or accounting access all require view access first.
// - Accounting mode is a fixed bundle (view-only + notes) that overrides add/edit/delete.
const reconcilePermissions = (perm) => {
  if (perm.isAccounting) {
    return { ...perm, canViewAll: true, canAdd: false, canEdit: false, canDelete: false };
  }
  return { ...perm, canViewAll: perm.canViewAll || perm.canEdit || perm.canDelete };
};

const emptyNewEmployee = {
  name: "",
  username: "",
  password: "",
  role: "employee",
  // Default permissions for a newly created employee: can only see and add
  // their own tickets, cannot edit or delete anything, and is not an accounting account.
  canViewAll: false,
  canAdd: true,
  canEdit: false,
  canDelete: false,
  isAccounting: false,
  canManageCompanies: false,
  isOwner: false,
  sections: { ...DEFAULT_SECTIONS },
  // Per-section view-all/edit/delete overrides; empty until customized per section, in
  // which case each section falls back to the account-wide toggles above.
  sectionPerms: {},
};

// Full-screen modal for editing one employee's grade and detailed permissions. Centered
// over the whole page (not nested inside the scrollable/clipped table), so it's always
// fully visible and easy to use — this is the one place permissions for an existing
// employee are changed. Closes itself if the employee record disappears (e.g. deleted
// from another tab) or is promoted to a main account (which no longer uses these toggles).
const EmployeePermissionsModal = ({ emp, onClose, onSetRole, onSetPermission, onSetSection, onSetSectionPerm, onSave, onDelete }) => {
  // Name/username/password are edited right here, not in the employee table — the
  // table's name cell is a plain, non-editable label that only opens this modal.
  const [draft, setDraft] = useState({ name: "", username: "", password: "" });
  const [showPassword, setShowPassword] = useState(false);
  const [saveError, setSaveError] = useState("");
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    if (emp) {
      setDraft({ name: emp.name, username: emp.username, password: emp.password });
      setShowPassword(false);
      setSaveError("");
    }
  }, [emp && emp.username]);

  if (!emp) return null;
  const sections = employeeSections(emp);

  const handleSaveDetails = async () => {
    setSaving(true);
    const err = await onSave({
      name: draft.name.trim(),
      username: draft.username.trim(),
      password: draft.password,
    });
    setSaving(false);
    setSaveError(err || "");
  };

  return (
    <div className="fixed inset-0 bg-stone-900/40 flex items-center justify-center p-4 z-50" onClick={onClose}>
      <div
        className="bg-white rounded-2xl border border-stone-200 p-5 w-full max-w-sm max-h-[90vh] overflow-y-auto"
        onClick={(ev) => ev.stopPropagation()}
      >
        <div className="flex items-center justify-between mb-4">
          <h3 className="font-semibold text-stone-900">Permissions</h3>
          <button onClick={onClose} className="text-stone-400 hover:text-stone-700 p-1">
            <X size={16} />
          </button>
        </div>

        <p className="text-xs text-stone-500 mb-1">Account details</p>
        <div className="border border-stone-200 rounded-xl p-3 mb-4 space-y-2">
          <div>
            <label className="text-xs text-stone-500 block mb-1">Full name</label>
            <input
              className="w-full border border-stone-300 rounded-xl px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
              value={draft.name}
              onChange={(ev) => setDraft({ ...draft, name: ev.target.value })}
            />
          </div>
          <div>
            <label className="text-xs text-stone-500 block mb-1">Username</label>
            <input
              className="w-full border border-stone-300 rounded-xl px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
              value={draft.username}
              onChange={(ev) => setDraft({ ...draft, username: ev.target.value })}
            />
          </div>
          <div>
            <label className="text-xs text-stone-500 block mb-1">Password</label>
            <div className="relative">
              <input
                type={showPassword ? "text" : "password"}
                className="w-full border border-stone-300 rounded-xl pl-2 pr-8 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={draft.password}
                onChange={(ev) => setDraft({ ...draft, password: ev.target.value })}
              />
              <button
                type="button"
                onClick={() => setShowPassword(!showPassword)}
                className="absolute right-2 top-1/2 -translate-y-1/2 text-stone-400"
              >
                {showPassword ? <EyeOff size={14} /> : <Eye size={14} />}
              </button>
            </div>
          </div>
          {saveError && <p className="text-xs text-red-600">{saveError}</p>}
          <button
            type="button"
            onClick={handleSaveDetails}
            disabled={saving}
            className="w-full border border-stone-300 hover:bg-stone-50 text-stone-700 text-sm font-semibold rounded-xl px-4 py-1.5 transition-colors disabled:opacity-50"
          >
            Save details
          </button>
        </div>

        {/* The clickable Grade grid was removed from this screen — grades are now
            assigned when the employee is first added, from the three tier dropdowns on
            the Add employee page. This just shows the grade as read-only info; every
            individual permission below can still be changed by hand regardless of it. */}
        <p className="text-xs text-stone-500 mb-1">Grade</p>
        <p className="text-sm font-medium text-stone-700 mb-4">{roleLabel(emp.role)}</p>

        <p className="text-xs text-stone-500 mb-1">Individual permissions</p>
        <div className="border border-stone-200 rounded-xl px-3 divide-y divide-stone-100 mb-4">
          <ToggleSwitch
            label="Accounting mode"
            description="View all tickets, but the only edit allowed is the Notes field"
            checked={emp.isAccounting}
            onChange={(v) => onSetPermission("isAccounting", v)}
          />
          <ToggleSwitch
            label="Manage companies"
            description="Add, edit, or remove saved company records"
            checked={emp.canManageCompanies}
            onChange={(v) => onSetPermission("canManageCompanies", v)}
          />
          <ToggleSwitch
            label="Owner access"
            description="Admin-level access — manage employees, backup/restore — everything except the License panel"
            checked={emp.isOwner}
            onChange={(v) => onSetPermission("isOwner", v)}
          />
        </div>

        <p className="text-xs text-stone-500 mb-1">Section access &amp; permissions</p>
        <p className="text-[11px] text-stone-400 mb-2">
          Turn a section on or off, then set exactly what they can do inside it — View all services, Edit, and Delete are each independent per section.
        </p>
        <div className="space-y-2">
          {SECTION_OPTIONS.map((s) => {
            const sectionOn = !!sections[s.value];
            const perm = employeeSectionPerm(emp, s.value);
            const hasOwnership = SECTIONS_WITH_OWNERSHIP.includes(s.value);
            return (
              <div key={s.value} className="border border-stone-200 rounded-xl px-3 divide-y divide-stone-100 overflow-hidden">
                <ToggleSwitch
                  label={s.label}
                  checked={sectionOn}
                  onChange={(v) => onSetSection(s.value, v)}
                />
                {sectionOn && (
                  <div className={emp.isAccounting ? "opacity-50 pointer-events-none" : ""}>
                    {hasOwnership && (
                      <ToggleSwitch
                        label="View all services"
                        description={`See every employee's ${s.label.toLowerCase()} records, not just their own`}
                        checked={perm.canViewAll}
                        onChange={(v) => onSetSectionPerm(s.value, "canViewAll", v)}
                      />
                    )}
                    <ToggleSwitch
                      label="Edit"
                      description={`Edit ${s.label.toLowerCase()} records they can see`}
                      checked={perm.canEdit}
                      onChange={(v) => onSetSectionPerm(s.value, "canEdit", v)}
                    />
                    <ToggleSwitch
                      label="Delete"
                      description={`Permanently remove ${s.label.toLowerCase()} records they can see`}
                      checked={perm.canDelete}
                      onChange={(v) => onSetSectionPerm(s.value, "canDelete", v)}
                    />
                  </div>
                )}
              </div>
            );
          })}
        </div>

        {onDelete && (
          <div className="flex gap-2 mt-4">
            <button
              type="button"
              onClick={onDelete}
              className="flex-1 flex items-center justify-center gap-1.5 border border-red-200 hover:bg-red-50 text-red-600 text-sm font-semibold rounded-xl px-4 py-2 transition-colors"
            >
              <Trash2 size={14} /> Delete
            </button>
          </div>
        )}

        <button
          onClick={onClose}
          className="mt-2 w-full bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 shadow-sm shadow-teal-800/30 transition-colors"
        >
          Done
        </button>
      </div>
    </div>
  );
};

// A small reusable on/off switch used throughout the permissions UI.
const ToggleSwitch = ({ checked, onChange, disabled, label, description }) => (
  <label className={`flex items-start justify-between gap-3 py-1.5 ${disabled ? "opacity-50" : ""}`}>
    <span>
      <span className="text-sm text-stone-700 font-medium block">{label}</span>
      {description && <span className="text-[11px] text-stone-400 block">{description}</span>}
    </span>
    <button
      type="button"
      disabled={disabled}
      onClick={() => !disabled && onChange(!checked)}
      className={`shrink-0 mt-0.5 relative inline-flex h-5 w-9 items-center rounded-full transition-colors ${
        checked ? "bg-teal-700" : "bg-stone-300"
      } ${disabled ? "cursor-not-allowed" : "cursor-pointer"}`}
    >
      <span
        className="inline-block h-3.5 w-3.5 transform rounded-full bg-white transition-transform"
        style={{ transform: checked ? "translateX(18px)" : "translateX(2px)" }}
      />
    </button>
  </label>
);

// IATA 3-digit airline accounting/ticketing prefix codes — the first 3 digits of a
// standard e-ticket number identify the issuing airline. Used to link the ticket
// number prefix with the Airline field automatically in both directions.
const AIRLINE_CODES = [
  { code: "001", iata: "AA", name: "American Airlines" }, { code: "006", iata: "DL", name: "Delta Air Lines" },
  { code: "014", iata: "AC", name: "Air Canada" }, { code: "016", iata: "UA", name: "United Airlines" },
  { code: "020", iata: "SU", name: "Aeroflot" }, { code: "022", iata: "DE", name: "Condor" },
  { code: "027", iata: "AS", name: "Alaska Airlines" }, { code: "030", iata: "VY", name: "Vueling" },
  { code: "044", iata: "AR", name: "Aerolineas Argentinas" }, { code: "045", iata: "LA", name: "LATAM Airlines" },
  { code: "050", iata: "OA", name: "Olympic Air" }, { code: "053", iata: "EI", name: "Aer Lingus" },
  { code: "055", iata: "AZ", name: "ITA Airways" }, { code: "057", iata: "AF", name: "Air France" },
  { code: "065", iata: "SV", name: "Saudia" }, { code: "071", iata: "ET", name: "Ethiopian Airlines" },
  { code: "072", iata: "GF", name: "Gulf Air" }, { code: "074", iata: "KL", name: "KLM Royal Dutch Airlines" },
  { code: "075", iata: "IB", name: "Iberia" }, { code: "076", iata: "ME", name: "Middle East Airlines" },
  { code: "077", iata: "MS", name: "EgyptAir" }, { code: "079", iata: "PR", name: "Philippine Airlines" },
  { code: "080", iata: "LO", name: "LOT Polish Airlines" }, { code: "081", iata: "QF", name: "Qantas" },
  { code: "082", iata: "SN", name: "Brussels Airlines" }, { code: "085", iata: "4Y", name: "Discover Airlines" },
  { code: "086", iata: "NZ", name: "Air New Zealand" }, { code: "087", iata: "DT", name: "TAAG Angola Airlines" },
  { code: "098", iata: "AI", name: "Air India" }, { code: "101", iata: "EN", name: "Air Dolomiti" },
  { code: "104", iata: "EW", name: "Eurowings" }, { code: "105", iata: "AY", name: "Finnair" },
  { code: "108", iata: "FI", name: "Icelandair" }, { code: "114", iata: "LY", name: "El Al" },
  { code: "115", iata: "JU", name: "Air Serbia" }, { code: "117", iata: "SK", name: "Scandinavian Airlines" },
  { code: "124", iata: "AH", name: "Air Algerie" }, { code: "125", iata: "BA", name: "British Airways" },
  { code: "126", iata: "GA", name: "Garuda Indonesia" }, { code: "127", iata: "G3", name: "Gol Transportes Aereos" },
  { code: "131", iata: "JL", name: "Japan Airlines" }, { code: "134", iata: "AV", name: "Avianca" },
  { code: "139", iata: "AM", name: "Aeromexico" }, { code: "147", iata: "AT", name: "Royal Air Maroc" },
  { code: "157", iata: "QR", name: "Qatar Airways" }, { code: "160", iata: "CX", name: "Cathay Pacific" },
  { code: "176", iata: "EK", name: "Emirates" }, { code: "180", iata: "KE", name: "Korean Air" },
  { code: "205", iata: "NH", name: "All Nippon Airways" }, { code: "217", iata: "TG", name: "Thai Airways International" },
  { code: "220", iata: "LH", name: "Lufthansa" }, { code: "230", iata: "CM", name: "Copa Airlines" },
  { code: "232", iata: "MH", name: "Malaysia Airlines" }, { code: "235", iata: "TK", name: "Turkish Airlines" },
  { code: "257", iata: "OS", name: "Austrian Airlines" }, { code: "279", iata: "B6", name: "JetBlue Airways" },
  { code: "281", iata: "RO", name: "TAROM" }, { code: "282", iata: "TP", name: "TAP Air Portugal" },
  { code: "297", iata: "CI", name: "China Airlines" }, { code: "312", iata: "6E", name: "IndiGo" },
  { code: "324", iata: "SC", name: "Shandong Airlines" }, { code: "328", iata: "DY", name: "Norwegian Air Shuttle" },
  { code: "390", iata: "A3", name: "Aegean Airlines" }, { code: "427", iata: "TX", name: "Air Caraibes" },
  { code: "465", iata: "KC", name: "Air Astana" }, { code: "479", iata: "ZH", name: "Shenzhen Airlines" },
  { code: "512", iata: "RJ", name: "Royal Jordanian" }, { code: "514", iata: "G9", name: "Air Arabia" },
  { code: "605", iata: "H2", name: "Sky Airline" }, { code: "607", iata: "EY", name: "Etihad Airways" },
  { code: "618", iata: "SQ", name: "Singapore Airlines" }, { code: "623", iata: "FB", name: "Bulgaria Air" },
  { code: "643", iata: "KM", name: "Air Malta" }, { code: "649", iata: "TS", name: "Air Transat" },
  { code: "657", iata: "BT", name: "Air Baltic" }, { code: "668", iata: "TR", name: "Scoot" },
  { code: "695", iata: "BR", name: "EVA Air" }, { code: "706", iata: "KQ", name: "Kenya Airways" },
  { code: "724", iata: "LX", name: "Swiss International Air Lines" }, { code: "731", iata: "MF", name: "Xiamen Airlines" },
  { code: "738", iata: "VN", name: "Vietnam Airlines" }, { code: "755", iata: "UX", name: "Air Europa" },
  { code: "774", iata: "FM", name: "Shanghai Airlines" }, { code: "781", iata: "MU", name: "China Eastern Airlines" },
  { code: "784", iata: "CZ", name: "China Southern Airlines" }, { code: "795", iata: "VA", name: "Virgin Australia" },
  { code: "821", iata: "NO", name: "Neos" }, { code: "831", iata: "OU", name: "Croatia Airlines" },
  { code: "838", iata: "WS", name: "WestJet" }, { code: "847", iata: "RX", name: "Riyadh Air" },
  { code: "876", iata: "3U", name: "Sichuan Airlines" }, { code: "880", iata: "HU", name: "Hainan Airlines" },
  { code: "900", iata: "F3", name: "flyadeal" }, { code: "932", iata: "VS", name: "Virgin Atlantic" },
  { code: "978", iata: "VJ", name: "VietJet Air" }, { code: "999", iata: "CA", name: "Air China" },
];
const getAirlineCode = (name) => {
  const n = (name || "").trim().toUpperCase();
  if (!n) return null;
  const match = AIRLINE_CODES.find((a) => a.name.toUpperCase() === n);
  return match ? match.code : null;
};
const getAirlineByCode = (code) => {
  const match = AIRLINE_CODES.find((a) => a.code === code);
  return match ? match.iata : null;
};
// 2-letter IATA airline designator (e.g. "MS" for EgyptAir) — this is what gets
// typed/selected into the Airline field and stored on the ticket.
const getAirlineIata = (name) => {
  const n = (name || "").trim().toUpperCase();
  if (!n) return null;
  const match = AIRLINE_CODES.find((a) => a.name.toUpperCase() === n);
  return match ? match.iata : null;
};
// Reverse lookups from the 2-letter code: the 3-digit accounting/ticketing prefix
// (used to auto-fill the ticket number) and the full airline name (shown as a hint).
const getAirlineCodeByIata = (iata) => {
  const n = (iata || "").trim().toUpperCase();
  if (!n) return null;
  const match = AIRLINE_CODES.find((a) => a.iata === n);
  return match ? match.code : null;
};
const getAirlineNameByIata = (iata) => {
  const n = (iata || "").trim().toUpperCase();
  if (!n) return null;
  const match = AIRLINE_CODES.find((a) => a.iata === n);
  return match ? match.name : null;
};

// A reference list of major world airports (IATA code + city/country), offered as
// autocomplete suggestions on the From/To fields alongside previously typed values.
const AIRPORTS = [
  ["CAI", "Cairo, Egypt"], ["HRG", "Hurghada, Egypt"], ["SSH", "Sharm El Sheikh, Egypt"],
  ["LXR", "Luxor, Egypt"], ["ASW", "Aswan, Egypt"], ["ALY", "Borg El Arab, Alexandria, Egypt"],
  ["DXB", "Dubai, UAE"], ["AUH", "Abu Dhabi, UAE"], ["SHJ", "Sharjah, UAE"],
  ["DOH", "Doha, Qatar"], ["KWI", "Kuwait City, Kuwait"], ["RUH", "Riyadh, Saudi Arabia"],
  ["JED", "Jeddah, Saudi Arabia"], ["DMM", "Dammam, Saudi Arabia"], ["MED", "Medina, Saudi Arabia"],
  ["BAH", "Manama, Bahrain"], ["MCT", "Muscat, Oman"], ["AMM", "Amman, Jordan"],
  ["BEY", "Beirut, Lebanon"], ["DAM", "Damascus, Syria"], ["BGW", "Baghdad, Iraq"],
  ["BSR", "Basra, Iraq"], ["EBL", "Erbil, Iraq"], ["TLV", "Tel Aviv, Israel"],
  ["CMN", "Casablanca, Morocco"], ["RAK", "Marrakesh, Morocco"], ["ALG", "Algiers, Algeria"],
  ["TUN", "Tunis, Tunisia"], ["TIP", "Tripoli, Libya"], ["KRT", "Khartoum, Sudan"],
  ["ADD", "Addis Ababa, Ethiopia"], ["NBO", "Nairobi, Kenya"], ["JNB", "Johannesburg, South Africa"],
  ["CPT", "Cape Town, South Africa"], ["LOS", "Lagos, Nigeria"], ["ACC", "Accra, Ghana"],
  ["DKR", "Dakar, Senegal"], ["ABJ", "Abidjan, Ivory Coast"],
  ["LHR", "London Heathrow, UK"], ["LGW", "London Gatwick, UK"], ["MAN", "Manchester, UK"],
  ["CDG", "Paris Charles de Gaulle, France"], ["ORY", "Paris Orly, France"],
  ["AMS", "Amsterdam, Netherlands"], ["FRA", "Frankfurt, Germany"], ["MUC", "Munich, Germany"],
  ["BER", "Berlin, Germany"], ["MAD", "Madrid, Spain"], ["BCN", "Barcelona, Spain"],
  ["FCO", "Rome, Italy"], ["MXP", "Milan, Italy"], ["IST", "Istanbul, Turkey"],
  ["SAW", "Istanbul Sabiha, Turkey"], ["ATH", "Athens, Greece"], ["ZRH", "Zurich, Switzerland"],
  ["GVA", "Geneva, Switzerland"], ["VIE", "Vienna, Austria"], ["BRU", "Brussels, Belgium"],
  ["CPH", "Copenhagen, Denmark"], ["OSL", "Oslo, Norway"], ["ARN", "Stockholm, Sweden"],
  ["HEL", "Helsinki, Finland"], ["DUB", "Dublin, Ireland"], ["LIS", "Lisbon, Portugal"],
  ["WAW", "Warsaw, Poland"], ["PRG", "Prague, Czech Republic"], ["BUD", "Budapest, Hungary"],
  ["OTP", "Bucharest, Romania"], ["SOF", "Sofia, Bulgaria"], ["BEG", "Belgrade, Serbia"],
  ["KEF", "Reykjavik, Iceland"], ["SVO", "Moscow, Russia"], ["LED", "St Petersburg, Russia"],
  ["DEL", "Delhi, India"], ["BOM", "Mumbai, India"], ["BLR", "Bangalore, India"],
  ["MAA", "Chennai, India"], ["HYD", "Hyderabad, India"], ["CCU", "Kolkata, India"],
  ["COK", "Kochi, India"], ["KHI", "Karachi, Pakistan"], ["LHE", "Lahore, Pakistan"],
  ["ISB", "Islamabad, Pakistan"], ["DAC", "Dhaka, Bangladesh"], ["CMB", "Colombo, Sri Lanka"],
  ["KTM", "Kathmandu, Nepal"], ["BKK", "Bangkok, Thailand"], ["HKT", "Phuket, Thailand"],
  ["SIN", "Singapore"], ["KUL", "Kuala Lumpur, Malaysia"], ["CGK", "Jakarta, Indonesia"],
  ["DPS", "Bali, Indonesia"], ["MNL", "Manila, Philippines"], ["HAN", "Hanoi, Vietnam"],
  ["SGN", "Ho Chi Minh City, Vietnam"], ["PNH", "Phnom Penh, Cambodia"], ["RGN", "Yangon, Myanmar"],
  ["HKG", "Hong Kong"], ["TPE", "Taipei, Taiwan"], ["ICN", "Seoul, South Korea"],
  ["NRT", "Tokyo Narita, Japan"], ["HND", "Tokyo Haneda, Japan"], ["KIX", "Osaka, Japan"],
  ["PEK", "Beijing, China"], ["PVG", "Shanghai, China"], ["CAN", "Guangzhou, China"],
  ["SZX", "Shenzhen, China"], ["ALA", "Almaty, Kazakhstan"], ["TAS", "Tashkent, Uzbekistan"],
  ["GYD", "Baku, Azerbaijan"], ["TBS", "Tbilisi, Georgia"], ["EVN", "Yerevan, Armenia"],
  ["JFK", "New York JFK, USA"], ["EWR", "Newark, USA"], ["LGA", "New York LaGuardia, USA"],
  ["LAX", "Los Angeles, USA"], ["ORD", "Chicago, USA"], ["MIA", "Miami, USA"],
  ["ATL", "Atlanta, USA"], ["DFW", "Dallas, USA"], ["SFO", "San Francisco, USA"],
  ["IAD", "Washington DC, USA"], ["BOS", "Boston, USA"], ["YYZ", "Toronto, Canada"],
  ["YVR", "Vancouver, Canada"], ["YUL", "Montreal, Canada"], ["MEX", "Mexico City, Mexico"],
  ["GRU", "Sao Paulo, Brazil"], ["GIG", "Rio de Janeiro, Brazil"], ["EZE", "Buenos Aires, Argentina"],
  ["SCL", "Santiago, Chile"], ["BOG", "Bogota, Colombia"], ["LIM", "Lima, Peru"],
  ["SYD", "Sydney, Australia"], ["MEL", "Melbourne, Australia"], ["BNE", "Brisbane, Australia"],
  ["PER", "Perth, Australia"], ["AKL", "Auckland, New Zealand"],

  // --- Additional Africa ---
  ["DAR", "Dar es Salaam, Tanzania"], ["ZNZ", "Zanzibar, Tanzania"], ["EBB", "Entebbe/Kampala, Uganda"],
  ["KGL", "Kigali, Rwanda"], ["BJM", "Bujumbura, Burundi"], ["LUN", "Lusaka, Zambia"],
  ["HRE", "Harare, Zimbabwe"], ["MPM", "Maputo, Mozambique"], ["WDH", "Windhoek, Namibia"],
  ["GBE", "Gaborone, Botswana"], ["TNR", "Antananarivo, Madagascar"], ["MRU", "Port Louis, Mauritius"],
  ["SEZ", "Mahe Island, Seychelles"], ["LAD", "Luanda, Angola"], ["FIH", "Kinshasa, DR Congo"],
  ["BZV", "Brazzaville, Republic of Congo"], ["LBV", "Libreville, Gabon"], ["DLA", "Douala, Cameroon"],
  ["NSI", "Yaounde, Cameroon"], ["NDJ", "N'Djamena, Chad"], ["NIM", "Niamey, Niger"],
  ["OUA", "Ouagadougou, Burkina Faso"], ["BKO", "Bamako, Mali"], ["COO", "Cotonou, Benin"],
  ["LFW", "Lome, Togo"], ["FNA", "Freetown, Sierra Leone"], ["ROB", "Monrovia, Liberia"],
  ["CKY", "Conakry, Guinea"], ["BJL", "Banjul, Gambia"], ["NKC", "Nouakchott, Mauritania"],
  ["PHC", "Port Harcourt, Nigeria"], ["ABV", "Abuja, Nigeria"], ["KAN", "Kano, Nigeria"],
  ["ASM", "Asmara, Eritrea"], ["JIB", "Djibouti City, Djibouti"], ["MGQ", "Mogadishu, Somalia"],
  ["HGA", "Hargeisa, Somaliland"], ["JUB", "Juba, South Sudan"],

  // --- Additional Middle East ---
  ["AAN", "Al Ain, UAE"], ["RKT", "Ras Al Khaimah, UAE"], ["NJF", "Najaf, Iraq"],

  // --- Additional Europe ---
  ["EDI", "Edinburgh, UK"], ["GLA", "Glasgow, UK"], ["BHX", "Birmingham, UK"], ["BRS", "Bristol, UK"],
  ["NCE", "Nice, France"], ["LYS", "Lyon, France"], ["MRS", "Marseille, France"], ["TLS", "Toulouse, France"],
  ["HAM", "Hamburg, Germany"], ["DUS", "Dusseldorf, Germany"], ["STR", "Stuttgart, Germany"], ["CGN", "Cologne, Germany"],
  ["NAP", "Naples, Italy"], ["VCE", "Venice, Italy"], ["BLQ", "Bologna, Italy"], ["TRN", "Turin, Italy"],
  ["PMI", "Palma de Mallorca, Spain"], ["AGP", "Malaga, Spain"], ["SVQ", "Seville, Spain"], ["VLC", "Valencia, Spain"],
  ["BIO", "Bilbao, Spain"], ["OPO", "Porto, Portugal"], ["FAO", "Faro, Portugal"], ["LUX", "Luxembourg City, Luxembourg"],
  ["KRK", "Krakow, Poland"], ["GDN", "Gdansk, Poland"], ["BTS", "Bratislava, Slovakia"], ["LJU", "Ljubljana, Slovenia"],
  ["ZAG", "Zagreb, Croatia"], ["SPU", "Split, Croatia"], ["DBV", "Dubrovnik, Croatia"], ["SJJ", "Sarajevo, Bosnia and Herzegovina"],
  ["SKP", "Skopje, North Macedonia"], ["TIA", "Tirana, Albania"], ["PRN", "Pristina, Kosovo"],
  ["HER", "Heraklion, Greece"], ["RHO", "Rhodes, Greece"], ["CFU", "Corfu, Greece"], ["JTR", "Santorini, Greece"],
  ["MLA", "Valletta, Malta"], ["LCA", "Larnaca, Cyprus"], ["PFO", "Paphos, Cyprus"],
  ["RIX", "Riga, Latvia"], ["VNO", "Vilnius, Lithuania"], ["TLL", "Tallinn, Estonia"], ["MSQ", "Minsk, Belarus"],
  ["KBP", "Kyiv, Ukraine"], ["ODS", "Odesa, Ukraine"], ["LWO", "Lviv, Ukraine"], ["KIV", "Chisinau, Moldova"],
  ["GOT", "Gothenburg, Sweden"], ["BGO", "Bergen, Norway"], ["TRD", "Trondheim, Norway"], ["AAL", "Aalborg, Denmark"],

  // --- Additional Asia ---
  ["PKX", "Beijing Daxing, China"], ["CTU", "Chengdu, China"], ["XIY", "Xi'an, China"], ["KMG", "Kunming, China"],
  ["WUH", "Wuhan, China"], ["NKG", "Nanjing, China"], ["TSN", "Tianjin, China"], ["HGH", "Hangzhou, China"],
  ["CSX", "Changsha, China"], ["URC", "Urumqi, China"], ["HAK", "Haikou, China"], ["SYX", "Sanya, China"],
  ["MFM", "Macau"], ["KHH", "Kaohsiung, Taiwan"], ["OKA", "Okinawa, Japan"], ["FUK", "Fukuoka, Japan"],
  ["CTS", "Sapporo, Japan"], ["NGO", "Nagoya, Japan"], ["GMP", "Seoul Gimpo, South Korea"], ["PUS", "Busan, South Korea"],
  ["CJU", "Jeju, South Korea"], ["UBN", "Ulaanbaatar, Mongolia"], ["VTE", "Vientiane, Laos"], ["LPQ", "Luang Prabang, Laos"],
  ["REP", "Siem Reap, Cambodia"], ["MDL", "Mandalay, Myanmar"], ["BWN", "Bandar Seri Begawan, Brunei"],
  ["CEB", "Cebu, Philippines"], ["DVO", "Davao, Philippines"], ["SUB", "Surabaya, Indonesia"], ["KNO", "Medan, Indonesia"],
  ["UPG", "Makassar, Indonesia"], ["PNQ", "Pune, India"], ["AMD", "Ahmedabad, India"], ["GOI", "Goa, India"],
  ["JAI", "Jaipur, India"], ["LKO", "Lucknow, India"], ["PAT", "Patna, India"], ["IXC", "Chandigarh, India"],
  ["TRV", "Thiruvananthapuram, India"], ["MLE", "Male, Maldives"],
  ["NQZ", "Astana, Kazakhstan"], ["FRU", "Bishkek, Kyrgyzstan"], ["DYU", "Dushanbe, Tajikistan"], ["ASB", "Ashgabat, Turkmenistan"],

  // --- Additional North America ---
  ["PHX", "Phoenix, USA"], ["DEN", "Denver, USA"], ["SEA", "Seattle, USA"], ["LAS", "Las Vegas, USA"],
  ["MSP", "Minneapolis, USA"], ["DTW", "Detroit, USA"], ["PHL", "Philadelphia, USA"], ["CLT", "Charlotte, USA"],
  ["HOU", "Houston Hobby, USA"], ["IAH", "Houston, USA"], ["SAN", "San Diego, USA"], ["TPA", "Tampa, USA"],
  ["MCO", "Orlando, USA"], ["FLL", "Fort Lauderdale, USA"], ["HNL", "Honolulu, USA"], ["ANC", "Anchorage, USA"],
  ["PDX", "Portland, USA"], ["AUS", "Austin, USA"], ["SLC", "Salt Lake City, USA"], ["STL", "St Louis, USA"],
  ["BWI", "Baltimore, USA"], ["DCA", "Washington Reagan, USA"], ["MSY", "New Orleans, USA"], ["OAK", "Oakland, USA"],
  ["YYC", "Calgary, Canada"], ["YEG", "Edmonton, Canada"], ["YOW", "Ottawa, Canada"], ["YHZ", "Halifax, Canada"],
  ["YWG", "Winnipeg, Canada"], ["GDL", "Guadalajara, Mexico"], ["MTY", "Monterrey, Mexico"], ["CUN", "Cancun, Mexico"],
  ["SJD", "Los Cabos, Mexico"], ["PVR", "Puerto Vallarta, Mexico"], ["TIJ", "Tijuana, Mexico"],

  // --- Central America & Caribbean ---
  ["GUA", "Guatemala City, Guatemala"], ["SAL", "San Salvador, El Salvador"], ["TGU", "Tegucigalpa, Honduras"],
  ["MGA", "Managua, Nicaragua"], ["SJO", "San Jose, Costa Rica"], ["PTY", "Panama City, Panama"],
  ["HAV", "Havana, Cuba"], ["SDQ", "Santo Domingo, Dominican Republic"], ["PUJ", "Punta Cana, Dominican Republic"],
  ["PAP", "Port-au-Prince, Haiti"], ["SJU", "San Juan, Puerto Rico"], ["MBJ", "Montego Bay, Jamaica"],
  ["KIN", "Kingston, Jamaica"], ["NAS", "Nassau, Bahamas"], ["BGI", "Bridgetown, Barbados"],
  ["POS", "Port of Spain, Trinidad and Tobago"], ["ANU", "St John's, Antigua"], ["BZE", "Belize City, Belize"],
  ["CUR", "Willemstad, Curacao"], ["AUA", "Oranjestad, Aruba"],

  // --- Additional South America ---
  ["BSB", "Brasilia, Brazil"], ["CNF", "Belo Horizonte, Brazil"], ["SSA", "Salvador, Brazil"],
  ["REC", "Recife, Brazil"], ["FOR", "Fortaleza, Brazil"], ["MAO", "Manaus, Brazil"],
  ["POA", "Porto Alegre, Brazil"], ["CWB", "Curitiba, Brazil"], ["MDZ", "Mendoza, Argentina"],
  ["COR", "Cordoba, Argentina"], ["USH", "Ushuaia, Argentina"], ["MVD", "Montevideo, Uruguay"],
  ["ASU", "Asuncion, Paraguay"], ["VVI", "Santa Cruz, Bolivia"], ["LPB", "La Paz, Bolivia"],
  ["UIO", "Quito, Ecuador"], ["GYE", "Guayaquil, Ecuador"], ["CTG", "Cartagena, Colombia"],
  ["MDE", "Medellin, Colombia"], ["CLO", "Cali, Colombia"], ["CCS", "Caracas, Venezuela"],
  ["GEO", "Georgetown, Guyana"], ["PBM", "Paramaribo, Suriname"],

  // --- Additional Oceania ---
  ["ADL", "Adelaide, Australia"], ["CNS", "Cairns, Australia"], ["OOL", "Gold Coast, Australia"],
  ["DRW", "Darwin, Australia"], ["HBA", "Hobart, Australia"], ["CHC", "Christchurch, New Zealand"],
  ["ZQN", "Queenstown, New Zealand"], ["WLG", "Wellington, New Zealand"], ["NAN", "Nadi, Fiji"],
  ["POM", "Port Moresby, Papua New Guinea"], ["NOU", "Noumea, New Caledonia"], ["PPT", "Papeete, Tahiti"],
  ["APW", "Apia, Samoa"], ["TBU", "Nuku'alofa, Tonga"], ["GUM", "Guam"], ["SPN", "Saipan"],
].map(([code, place]) => `${code} - ${place}`.toUpperCase());

// A small dropdown that lets the user tick any number of checkboxes instead of picking
// just one value. Used for every "Year / Month / Company / Employee / Supplier" style
// filter across the app so people can e.g. select two or three years at once.
// `options` can be an array of plain strings, or an array of { value, label } objects
// when the display text needs to differ from the underlying value (e.g. months).
function MultiSelectDropdown({ label, icon: Icon, options, selected, onChange, placeholder }) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef(null);

  useEffect(() => {
    if (!open) return;
    const handleClickOutside = (e) => {
      if (rootRef.current && !rootRef.current.contains(e.target)) setOpen(false);
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [open]);

  const normalized = options.map((opt) =>
    typeof opt === "object" && opt !== null ? opt : { value: opt, label: opt }
  );

  const toggleValue = (value) => {
    if (selected.includes(value)) {
      onChange(selected.filter((v) => v !== value));
    } else {
      onChange([...selected, value]);
    }
  };

  const selectedLabels = normalized
    .filter((opt) => selected.includes(opt.value))
    .map((opt) => opt.label);

  const displayText =
    selectedLabels.length === 0
      ? placeholder
      : selectedLabels.length === 1
      ? selectedLabels[0]
      : `${selectedLabels.length} selected`;

  return (
    <div className="relative" ref={rootRef}>
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className={`w-auto max-w-[160px] flex items-center gap-1 border rounded-lg pl-7 pr-2 py-1.5 text-xs focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white appearance-none relative ${
          selected.length > 0 ? "border-teal-700 text-teal-800" : "border-stone-300 text-stone-700"
        }`}
      >
        {Icon && (
          <Icon size={13} className="absolute left-2 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none" />
        )}
        <span className="truncate">{displayText}</span>
        <ChevronDown size={12} className={`ml-auto shrink-0 transition-transform ${open ? "rotate-180" : ""}`} />
      </button>
      {open && (
        <div className="absolute z-20 mt-1 w-48 max-h-64 overflow-auto bg-white border border-stone-200 rounded-lg shadow-lg py-1">
          {normalized.length === 0 && (
            <div className="px-3 py-2 text-xs text-stone-400">No options</div>
          )}
          {normalized.map((opt) => (
            <label
              key={opt.value}
              className="flex items-center gap-2 px-3 py-1.5 text-xs hover:bg-stone-50 cursor-pointer"
            >
              <input
                type="checkbox"
                checked={selected.includes(opt.value)}
                onChange={() => toggleValue(opt.value)}
                className="accent-teal-700"
              />
              <span className="truncate">{opt.label}</span>
            </label>
          ))}
          {selected.length > 0 && (
            <button
              type="button"
              onClick={() => onChange([])}
              className="w-full text-left px-3 py-1.5 text-xs text-teal-700 hover:bg-stone-50 border-t border-stone-100 mt-1"
            >
              Clear {label}
            </button>
          )}
        </div>
      )}
    </div>
  );
}

export default function TicketsApp({ onChangeServer, currentServerUrl } = {}) {
  // Prevent the mouse/trackpad scroll wheel from changing the value of a focused
  // number input. Browsers normally let scrolling over a focused number field
  // bump its value up/down, which is easy to trigger by accident while scrolling
  // the page. Blurring the field on wheel (while it's focused) stops that, and
  // since blur only fires when the field actually has focus, normal page
  // scrolling elsewhere is completely unaffected.
  useEffect(() => {
    const handleWheelOnNumberInput = (e) => {
      const active = document.activeElement;
      if (active && active.tagName === "INPUT" && active.type === "number") {
        active.blur();
      }
    };
    document.addEventListener("wheel", handleWheelOnNumberInput, { passive: true });
    return () => document.removeEventListener("wheel", handleWheelOnNumberInput);
  }, []);

  // Keep the browser tab title and icon in sync with the app's name/branding.
  useEffect(() => {
    document.title = "Travel Agency Manager";
    const faviconSvg =
      '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">' +
      '<rect width="24" height="24" rx="5" fill="#115e59"/>' +
      '<g transform="rotate(45 12 12)">' +
      '<path d="M21,16V14L13,9V3.5C13,2.67 12.33,2 11.5,2C10.67,2 10,2.67 10,3.5V9L2,14V16L10,13.5V19L7.5,20.5V22L11.5,21L15.5,22V20.5L13,19V13.5L21,16Z" fill="#ffffff"/>' +
      "</g></svg>";
    const href = "data:image/svg+xml," + encodeURIComponent(faviconSvg);
    let link = document.querySelector("link[rel~='icon']");
    if (!link) {
      link = document.createElement("link");
      link.rel = "icon";
      document.head.appendChild(link);
    }
    link.type = "image/svg+xml";
    link.href = href;
  }, []);

  // ---------- License / activation ----------
  // Stored centrally (shared storage) so activation applies to every employee,
  // not just the browser it was entered on. null = not loaded from storage yet.
  const [licenseRecord, setLicenseRecord] = useState(null); // { code, expiresAt } | null
  const [licenseLoaded, setLicenseLoaded] = useState(false);
  const [showLicensePanel, setShowLicensePanel] = useState(false);
  const [licenseInput, setLicenseInput] = useState("");
  const [licenseError, setLicenseError] = useState("");
  const [licenseSaving, setLicenseSaving] = useState(false);

  const licenseCheck = licenseRecord ? checkLicenseCode(licenseRecord.code) : { valid: false };
  const isLicensed = licenseCheck.valid;

  const handleActivateLicense = async () => {
    const result = checkLicenseCode(licenseInput);
    if (!result.valid) {
      setLicenseError(result.reason);
      return;
    }
    setLicenseError("");
    setLicenseSaving(true);
    try {
      await window.storage.set(
        "tickets:license",
        JSON.stringify({ code: result.code, expiresAt: result.expiresAt || null, activatedAt: Date.now() }),
        true
      );
      setLicenseRecord({ code: result.code, expiresAt: result.expiresAt || null });
      setLicenseInput("");
      setShowLicensePanel(false);
    } catch (e) {
      setLicenseError("Couldn't save the activation code — please try again.");
    } finally {
      setLicenseSaving(false);
    }
  };

  const [tickets, setTickets] = useState([]);
  const [employees, setEmployees] = useState(null); // null = not loaded yet
  const [currentUser, setCurrentUser] = useState(null); // { username, name, isAdmin }
  const [loading, setLoading] = useState(true);

  // Presence: which employees are currently connected (main account only)
  const [presenceMap, setPresenceMap] = useState({}); // username -> last-seen timestamp
  const [showOnlineList, setShowOnlineList] = useState(false);
  const [restoreError, setRestoreError] = useState("");
  const [restoreSuccess, setRestoreSuccess] = useState("");
  const fileInputRef = useRef(null);
  // Timestamp of when the current session started (login or restored on page load).
  // A remote force-sign-out (see handleForceSignOut) is only honored if it happened
  // AFTER this moment, so it can't retroactively sign someone out of a brand new session.
  const sessionStartedAtRef = useRef(0);
  // window.confirm doesn't work in this sandboxed preview, so confirmations use this
  // in-app dialog instead: { message, onConfirm } while open, null while hidden.
  const [confirmDialog, setConfirmDialog] = useState(null);
  const requestConfirm = (message, onConfirm) => setConfirmDialog({ message, onConfirm });

  // In-app print preview: { title, html } while open, null while closed. Printing renders
  // the receipt into a hidden iframe inside this popup instead of opening a separate
  // browser tab/window, so it can't be blocked and always stays part of the app.
  const [printPreview, setPrintPreview] = useState(null);
  const printIframeRef = useRef(null);

  // ---------- Employee-to-employee requests ----------
  // Lets any employee ask another employee for something (e.g. "please check this
  // customer's file"). Stored centrally like tickets/employees so it's visible to
  // everyone, and picked up by the same polling loop that already keeps other shared
  // data in sync — that's also how a teammate's screen finds out about a new request
  // addressed to them without a manual refresh.
  const [requests, setRequests] = useState([]);
  const [showRequestsPanel, setShowRequestsPanel] = useState(false);
  const [newRequestTo, setNewRequestTo] = useState("");
  const [newRequestMessage, setNewRequestMessage] = useState("");
  const [requestSendError, setRequestSendError] = useState("");
  // The single incoming request currently shown as a popup/notification, or null.
  // Only ever holds requests addressed to the signed-in account.
  const [incomingRequestPopup, setIncomingRequestPopup] = useState(null);
  // Ids of requests already shown as a popup this session, so the same one never pops
  // up twice (e.g. after the next poll re-fetches the same still-pending request).
  const seenRequestIdsRef = useRef(new Set());

  const [loginUsername, setLoginUsername] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  // Basic brute-force throttling: after 5 failed attempts in a row, lock the login form
  // for 30 seconds. This is in-memory only (resets on page reload), so it's friction
  // against casual/automated guessing, not a real security boundary.
  const [loginFailCount, setLoginFailCount] = useState(0);
  const [loginLockUntil, setLoginLockUntil] = useState(0);

  const [setupName, setSetupName] = useState("");
  const [setupUsername, setSetupUsername] = useState("");
  const [setupPassword, setSetupPassword] = useState("");

  const [showManage, setShowManage] = useState(false);
  const [newEmployee, setNewEmployee] = useState(emptyNewEmployee);
  const [newEmployeeGradeOpen, setNewEmployeeGradeOpen] = useState(null); // which of the three Grade dropdowns is open on the Add employee page: "manager" | "supervisor" | "employee" | null
  const [openPermissionsFor, setOpenPermissionsFor] = useState(null); // username, or null if closed
  const [manageError, setManageError] = useState("");
  const [editingUsername, setEditingUsername] = useState(null);
  const [editDraft, setEditDraft] = useState({ name: "", username: "", password: "" });
  const [editShowPassword, setEditShowPassword] = useState(false);

  const [showManageCompanies, setShowManageCompanies] = useState(false);
  const [showCompaniesList, setShowCompaniesList] = useState(false);
  const [newCompanyDraft, setNewCompanyDraft] = useState(emptyCompanyDraft);
  const [editingCompanyName, setEditingCompanyName] = useState(null);
  const [companyError, setCompanyError] = useState("");

  const [showChangePassword, setShowChangePassword] = useState(false);
  const [currentPasswordInput, setCurrentPasswordInput] = useState("");
  const [newPasswordInput, setNewPasswordInput] = useState("");
  const [confirmPasswordInput, setConfirmPasswordInput] = useState("");
  const [passwordError, setPasswordError] = useState("");
  const [passwordSuccess, setPasswordSuccess] = useState("");

  const [form, setForm] = useState(getEmptyForm);
  // Whether the Supplier field is in "type your own name" mode (chosen via the Other option).
  const [supplierOther, setSupplierOther] = useState(false);

  // ---------- Hotels ----------
  const [hotelBookings, setHotelBookings] = useState([]);
  const [hotelForm, setHotelForm] = useState(getEmptyHotelForm);
  const [hotelError, setHotelError] = useState("");
  const [hotelEditingId, setHotelEditingId] = useState(null);
  // The hotel booking currently shown in the read-only details modal (null = closed).
  const [viewingHotelBooking, setViewingHotelBooking] = useState(null);
  // Whether the "Add supplier" / "Add hotel name" panels at the top of the Hotels
  // page are currently open, plus the text typed into each panel's input.
  const [showAddSupplierPanel, setShowAddSupplierPanel] = useState(false);
  const [showAddHotelNamePanel, setShowAddHotelNamePanel] = useState(false);
  const [newSupplierDraft, setNewSupplierDraft] = useState("");
  const [newHotelNameDraft, setNewHotelNameDraft] = useState("");
  // Whether the Hotel name / Supplier fields on the booking form are in "type your
  // own name" mode, same pattern as supplierOther for flight tickets above.
  const [hotelSupplierOther, setHotelSupplierOther] = useState(false);
  const [hotelNameOther, setHotelNameOther] = useState(false);

  // ---------- Visa ----------
  const [visaBookings, setVisaBookings] = useState([]);
  const [visaForm, setVisaForm] = useState(getEmptyVisaForm);
  const [visaError, setVisaError] = useState("");
  const [visaEditingId, setVisaEditingId] = useState(null);
  // The visa booking currently shown in the read-only details modal (null = closed).
  const [viewingVisaBooking, setViewingVisaBooking] = useState(null);
  // Whether the Supplier field on the visa booking form is in "type your own name" mode,
  // same pattern as supplierOther / hotelSupplierOther above.
  const [visaSupplierOther, setVisaSupplierOther] = useState(false);
  // Whether the Visa page's own "Add supplier" panel is open, plus its draft text —
  // kept separate from the Hotels/Flights supplier panels so each section's suppliers
  // are independent lists.
  const [showAddVisaSupplierPanel, setShowAddVisaSupplierPanel] = useState(false);
  const [newVisaSupplierDraft, setNewVisaSupplierDraft] = useState("");

  // ---------- Transfers (Cars) ----------
  const [carBookings, setCarBookings] = useState([]);
  const [carForm, setCarForm] = useState(getEmptyCarForm);
  const [carError, setCarError] = useState("");
  const [carEditingId, setCarEditingId] = useState(null);
  // The transfer booking currently shown in the read-only details modal (null = closed).
  const [viewingCarBooking, setViewingCarBooking] = useState(null);
  // Whether the "other" free-text supplier field is shown instead of the dropdown list —
  // same pattern as visaSupplierOther above.
  const [carSupplierOther, setCarSupplierOther] = useState(false);
  const [showAddCarSupplierPanel, setShowAddCarSupplierPanel] = useState(false);
  const [newCarSupplierDraft, setNewCarSupplierDraft] = useState("");

  // ---------- Files ----------
  // A "file" bundles together copies (snapshots) of records already entered under
  // Flights/Hotels/Visa/Transportation, so their prices can be gathered and reviewed
  // together without touching the original records — nothing here feeds back into the
  // totals shown in those other sections.
  const [files, setFiles] = useState([]);
  const [fileError, setFileError] = useState("");
  // Which file (by id) is currently open in the detail view; null = showing the list.
  const [openFileId, setOpenFileId] = useState(null);
  // Whether the currently open file's services list is in "edit" mode (showing the Add
  // service button and each item's delete/trash icon). Off by default so the file detail
  // view opens as a clean read-only summary; toggled on with the "Edit services" button.
  const [editingFileServices, setEditingFileServices] = useState(false);
  // Whether the "add a copy from a service" picker is open, and which service tab it's on.
  const [showFilePicker, setShowFilePicker] = useState(false);
  const [filePickerTab, setFilePickerTab] = useState("flights");
  // A file being newly created but not yet confirmed/saved: null while not creating,
  // otherwise { company, notes, createdAt, items }. Lives only in local state — nothing is
  // written to the files table until the "Add file" (confirm) button is pressed. Services
  // can still be pulled in via "Add services" while in this draft state.
  const [draftFile, setDraftFile] = useState(null);
  // Set when "copy to a file" is clicked from the Flights/Hotels/Visa tables directly —
  // { type: 'flights'|'hotels'|'visa', record } — opens a modal asking which file (by
  // its serial number) to drop the copy into.
  const [copyPickerSource, setCopyPickerSource] = useState(null);
  // USD -> EGP exchange rate, used to also show a USD booking's value in EGP.
  // Entered by hand (no CBE API is publicly reachable from the browser), and saved so
  // everyone signed in sees today's rate without re-typing it.
  const [usdToEgpRate, setUsdToEgpRate] = useState(null);
  const [usdToEgpRateDate, setUsdToEgpRateDate] = useState("");

  // IATA balance tracker (Flights section): a running balance saved to shared storage,
  // and a separate "issued ticket value" box — entering an amount there deducts it from
  // the balance (see applyIataTicketValue below), so everyone signed in sees the same
  // running balance without each of them having to do the subtraction by hand.
  // iataBalanceLoaded stays false until the saved balance has actually been fetched —
  // deductions are blocked until then, so a deduction typed before the fetch resolves
  // can never be computed against an unloaded `null` and wipe out the real balance.
  const [iataBalance, setIataBalance] = useState(null);
  const [iataBalanceLoaded, setIataBalanceLoaded] = useState(false);
  const [iataTicketValueInput, setIataTicketValueInput] = useState("");
  // History of today's deductions from the IATA balance only — starts empty again at
  // the first deduction of each new day (see recordIataDeduction below). Kept in its own
  // shared-storage key, entirely separate from tickets/customers/accounts, so these two
  // fields never feed into any other totals. Viewed via the "History" button, which
  // opens it in a separate popup (showIataHistory below).
  const [iataHistory, setIataHistory] = useState({ date: "", deductions: [] });
  const [showIataHistory, setShowIataHistory] = useState(false);

  // ---------- Accounts (accounting module) ----------
  const [expenses, setExpenses] = useState([]);
  const [supplierPayments, setSupplierPayments] = useState([]);
  const [customerPayments, setCustomerPayments] = useState([]);
  const [treasuryAccounts, setTreasuryAccounts] = useState([]);
  const [treasuryEntries, setTreasuryEntries] = useState([]);
  // Which sub-tab of the Accounts section is showing.
  const [accountsTab, setAccountsTab] = useState("overview");
  const [accountsError, setAccountsError] = useState("");
  // Display-language toggle for the Accounts section only (Arabic/English). Stored
  // values (expense categories, entry categories, etc.) always stay in Arabic — this
  // only controls what's shown on screen.
  const [accountsLang, setAccountsLang] = useState("ar");
  const at = (key) => {
    const val = ACCOUNTS_I18N[accountsLang][key];
    return val === undefined ? key : val;
  };
  const acctCurrency = ACCOUNTS_I18N[accountsLang].currency;
  const expenseCategoryLabel = (cat) => (accountsLang === "en" ? (EXPENSE_CATEGORY_LABELS_EN[cat] || cat) : cat);
  const treasuryAccountTypeLabel = (val) => (accountsLang === "en" ? (TREASURY_ACCOUNT_TYPE_LABELS_EN[val] || val) : val);
  const treasuryEntryCategoryLabel = (cat) => (accountsLang === "en" ? (TREASURY_ENTRY_CATEGORY_LABELS_EN[cat] || cat) : cat);

  const [expenseForm, setExpenseForm] = useState(getEmptyExpenseForm);
  const [expenseEditingId, setExpenseEditingId] = useState(null);
  const [showExpenseForm, setShowExpenseForm] = useState(false);
  const [expenseCategoryFilter, setExpenseCategoryFilter] = useState("");

  const [treasuryForm, setTreasuryForm] = useState(getEmptyTreasuryAccountForm);
  const [showTreasuryAccountForm, setShowTreasuryAccountForm] = useState(false);
  const [treasuryAccountEditingId, setTreasuryAccountEditingId] = useState(null);
  const [treasuryEntryForm, setTreasuryEntryForm] = useState(getEmptyTreasuryEntryForm);
  const [showTreasuryEntryForm, setShowTreasuryEntryForm] = useState(false);
  const [treasuryFilterAccountId, setTreasuryFilterAccountId] = useState("");

  // Supplier/customer ledgers are derived, not stored — drilling into one just
  // remembers its name here so the detail panel knows which one to show.
  const [viewingSupplier, setViewingSupplier] = useState(null);
  const [viewingCustomer, setViewingCustomer] = useState(null);
  const [supplierQuery, setSupplierQuery] = useState("");
  const [customerQuery, setCustomerQuery] = useState("");
  const [supplierPaymentForm, setSupplierPaymentForm] = useState(getEmptySupplierPaymentForm);
  const [customerPaymentForm, setCustomerPaymentForm] = useState(getEmptyCustomerPaymentForm);

  // Date range for the Reports sub-tab: "today" | "month" | "custom".
  const [reportsRange, setReportsRange] = useState("month");
  const [reportsFrom, setReportsFrom] = useState("");
  const [reportsTo, setReportsTo] = useState("");

  const [query, setQuery] = useState("");
  const [error, setError] = useState("");
  // Clicking a ticket row opens a full detail view of that ticket (id stored here).
  const [viewingTicketId, setViewingTicketId] = useState(null);
  // Which row is currently highlighted/jumped-to in the main tickets table, identified
  // by a "ticket:<TICKETNUMBER>" or "refund:<TICKETNUMBER>" key (see data-row-key on
  // each <tr> below). Set when clicking an "Exchanged"/"Refunded" badge so the related
  // row — including an old, reissued ticket that's normally hidden from the table — can
  // be scrolled to and briefly highlighted.
  const [highlightedRowKey, setHighlightedRowKey] = useState(null);
  const highlightTimeoutRef = useRef(null);
  const jumpToRow = (key) => {
    if (!key) return;
    setHighlightedRowKey(key);
    if (highlightTimeoutRef.current) clearTimeout(highlightTimeoutRef.current);
    setTimeout(() => {
      const el = document.querySelector(`tr[data-row-key="${key}"]`);
      if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
    }, 60);
    highlightTimeoutRef.current = setTimeout(() => {
      setHighlightedRowKey((cur) => (cur === key ? null : cur));
    }, 4000);
  };
  const [notesDraft, setNotesDraft] = useState("");
  const [notesSaved, setNotesSaved] = useState(false);
  // Refund box in the main ticket form (next to Reissue): looks up existing tickets by
  // number and records a refund against each directly, independent of whichever ticket
  // the form itself is currently adding/editing. Supports refunding several tickets at
  // once — each row is its own ticket-number lookup plus its own amounts.
  const [refundBoxOpen, setRefundBoxOpen] = useState(false);
  const [refundRows, setRefundRows] = useState([{ number: "", airlineAmount: "", customerAmount: "", customerIndex: 0 }]);
  const [refundSaved, setRefundSaved] = useState(false);
  const [selectedMonth, setSelectedMonth] = useState([]);
  const [selectedYear, setSelectedYear] = useState([]);
  const [selectedCompany, setSelectedCompany] = useState([]);
  const [selectedEmployee, setSelectedEmployee] = useState([]);
  const [selectedSupplier, setSelectedSupplier] = useState([]);
  const [filtersOpen, setFiltersOpen] = useState(false);

  // Search + filter state for the Hotels, Visa, Transportation, and Files sections —
  // each section gets its own independent search box and filter set (mirroring the
  // Flights search/filters above), matched to the fields that section actually has.
  // Each filter below holds an array of selected values, so more than one option
  // (e.g. two years, or three employees) can be picked at once.
  const [hotelQuery, setHotelQuery] = useState("");
  const [hotelFiltersOpen, setHotelFiltersOpen] = useState(false);
  const [hotelSelectedYear, setHotelSelectedYear] = useState([]);
  const [hotelSelectedMonth, setHotelSelectedMonth] = useState([]);
  const [hotelSelectedEmployee, setHotelSelectedEmployee] = useState([]);
  const [hotelSelectedSupplier, setHotelSelectedSupplier] = useState([]);
  const [hotelSelectedHotelName, setHotelSelectedHotelName] = useState([]);

  const [visaQuery, setVisaQuery] = useState("");
  const [visaFiltersOpen, setVisaFiltersOpen] = useState(false);
  const [visaSelectedYear, setVisaSelectedYear] = useState([]);
  const [visaSelectedMonth, setVisaSelectedMonth] = useState([]);
  const [visaSelectedEmployee, setVisaSelectedEmployee] = useState([]);
  const [visaSelectedSupplier, setVisaSelectedSupplier] = useState([]);

  const [carQuery, setCarQuery] = useState("");
  const [carFiltersOpen, setCarFiltersOpen] = useState(false);
  const [carSelectedYear, setCarSelectedYear] = useState([]);
  const [carSelectedMonth, setCarSelectedMonth] = useState([]);
  const [carSelectedEmployee, setCarSelectedEmployee] = useState([]);
  const [carSelectedSupplier, setCarSelectedSupplier] = useState([]);

  const [fileQuery, setFileQuery] = useState("");
  const [fileFiltersOpen, setFileFiltersOpen] = useState(false);
  const [fileSelectedYear, setFileSelectedYear] = useState([]);
  const [fileSelectedCompany, setFileSelectedCompany] = useState([]);
  const [fileSelectedEmployee, setFileSelectedEmployee] = useState([]);

  // Every value ever entered (companies, customers, airlines, cities) is kept here so it
  // can be offered as an autocomplete suggestion later, even if the original ticket is deleted.
  const [suggestions, setSuggestions] = useState({ companies: [], customers: [], airlines: [], cities: [], suppliers: [], hotelNames: [], visaSuppliers: [], carSuppliers: [] });

  // Tracks whether the one-time "create the main account" step has ever been completed.
  // Once true, the first-run setup screen must never be shown again — even if the employee
  // list later becomes empty (e.g. accounts deleted, a bad restore) — so no one can
  // create a fresh, unauthenticated admin account after the app has already been set up.
  const [setupComplete, setSetupComplete] = useState(null); // null = not loaded yet

  // Top-level section switcher: "flights" holds all existing ticket functionality;
  // "hotels" and "cars" are placeholders for future sections.
  const [activeSection, setActiveSection] = useState("flights");

  // Remembers which section (flights/hotels/cars/files) this account was on, so a page
  // refresh returns to the same place instead of resetting to Flights. Skipped on the
  // very first render for a session, since that value was just restored from storage
  // above (or is the deliberate default) rather than a change the user made.
  const sectionHydratedRef = useRef(false);
  useEffect(() => {
    if (!currentUser) return;
    if (!sectionHydratedRef.current) {
      sectionHydratedRef.current = true;
      return;
    }
    window.storage.set(`tickets:lastSection:${currentUser.username}`, activeSection, false).catch(() => {});
  }, [activeSection, currentUser]);

  useEffect(() => {
    (async () => {
      try {
        const [ticketsRes, hotelsRes, visasRes, carsRes, filesRes, employeesRes, sessionRes, suggestionsRes, setupRes, licenseRes, requestsRes, expensesRes, supplierPaymentsRes, customerPaymentsRes, treasuryAccountsRes, treasuryEntriesRes] = await Promise.all([
          window.storage.get("tickets:list", true).catch(() => null),
          window.storage.get("tickets:hotels", true).catch(() => null),
          window.storage.get("tickets:visas", true).catch(() => null),
          window.storage.get("tickets:cars", true).catch(() => null),
          window.storage.get("tickets:files", true).catch(() => null),
          window.storage.get("tickets:employees", true).catch(() => null),
          window.storage.get("session:user", false).catch(() => null),
          window.storage.get("tickets:suggestions", true).catch(() => null),
          window.storage.get("tickets:setupComplete", true).catch(() => null),
          window.storage.get("tickets:license", true).catch(() => null),
          window.storage.get("tickets:requests", true).catch(() => null),
          window.storage.get("tickets:expenses", true).catch(() => null),
          window.storage.get("tickets:supplierPayments", true).catch(() => null),
          window.storage.get("tickets:customerPayments", true).catch(() => null),
          window.storage.get("tickets:treasuryAccounts", true).catch(() => null),
          window.storage.get("tickets:treasuryEntries", true).catch(() => null),
        ]);
        const ticketsData = ticketsRes && ticketsRes.value ? JSON.parse(ticketsRes.value) : [];
        const hotelsData = hotelsRes && hotelsRes.value ? JSON.parse(hotelsRes.value) : [];
        const visasData = visasRes && visasRes.value ? JSON.parse(visasRes.value) : [];
        const carsData = carsRes && carsRes.value ? JSON.parse(carsRes.value) : [];
        const filesData = filesRes && filesRes.value ? JSON.parse(filesRes.value) : [];
        const employeesData = employeesRes && employeesRes.value ? JSON.parse(employeesRes.value) : [];
        const requestsData = requestsRes && requestsRes.value ? JSON.parse(requestsRes.value) : [];
        setTickets(ticketsData);
        setHotelBookings(hotelsData);
        setVisaBookings(visasData);
        setCarBookings(carsData);
        setFiles(filesData);
        setEmployees(employeesData);
        setRequests(requestsData);
        setExpenses(expensesRes && expensesRes.value ? JSON.parse(expensesRes.value) : []);
        setSupplierPayments(supplierPaymentsRes && supplierPaymentsRes.value ? JSON.parse(supplierPaymentsRes.value) : []);
        setCustomerPayments(customerPaymentsRes && customerPaymentsRes.value ? JSON.parse(customerPaymentsRes.value) : []);
        setTreasuryAccounts(treasuryAccountsRes && treasuryAccountsRes.value ? JSON.parse(treasuryAccountsRes.value) : []);
        setTreasuryEntries(treasuryEntriesRes && treasuryEntriesRes.value ? JSON.parse(treasuryEntriesRes.value) : []);
        requestsData.forEach((r) => seenRequestIdsRef.current.add(r.id));
        if (licenseRes && licenseRes.value) {
          try {
            setLicenseRecord(JSON.parse(licenseRes.value));
          } catch (e) {
            setLicenseRecord(null);
          }
        }
        setLicenseLoaded(true);
        // If accounts already exist, the setup step has clearly already happened even if the
        // flag itself is missing (e.g. app used before this flag existed).
        setSetupComplete(!!(setupRes && setupRes.value === "true") || employeesData.length > 0);
        if (suggestionsRes && suggestionsRes.value) {
          try {
            const parsed = JSON.parse(suggestionsRes.value);
            setSuggestions({
              companies: parsed.companies || [],
              customers: [], // never restore saved customer names — this field must have no autocomplete history
              airlines: parsed.airlines || [],
              cities: parsed.cities || [],
              suppliers: parsed.suppliers || [],
              hotelNames: parsed.hotelNames || [],
              visaSuppliers: parsed.visaSuppliers || [],
              carSuppliers: parsed.carSuppliers || [],
            });
          } catch (e) {
            // ignore malformed suggestions data
          }
        }

        if (sessionRes && sessionRes.value) {
          const savedUsername = sessionRes.value;
          const match = employeesData.find((e) => e.username === savedUsername);
          if (match) {
            sessionStartedAtRef.current = Date.now();
            setCurrentUser({ username: match.username, name: match.name, isAdmin: !!match.isAdmin });
            try {
              const lastSectionRes = await window.storage.get(`tickets:lastSection:${match.username}`, false).catch(() => null);
              const lastSection = lastSectionRes && lastSectionRes.value;
              if (["flights", "hotels", "visa", "cars", "files"].includes(lastSection)) {
                setActiveSection(lastSection);
              }
            } catch (e) {
              // Best-effort; falls back to the default "flights" section
            }
          }
        }
      } catch (e) {
        setEmployees([]);
        setSetupComplete(false);
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  const LIVE_REFRESH_INTERVAL_MS = 5 * 1000;

  // Keeps tickets, employee accounts, and saved suggestions (companies/customers/
  // airlines/cities) in sync across everyone who's signed in, by periodically re-reading
  // the shared storage keys. window.storage has no push/subscribe API, so short polling
  // is the only way to reflect other users' changes without a manual page refresh.
  useEffect(() => {
    if (!currentUser) return;
    let cancelled = false;
    const loadCoreData = async () => {
      try {
        const [ticketsRes, hotelsRes, visasRes, carsRes, filesRes, employeesRes, suggestionsRes, licenseRes, requestsRes, expensesRes, supplierPaymentsRes, customerPaymentsRes, treasuryAccountsRes, treasuryEntriesRes] = await Promise.all([
          window.storage.get("tickets:list", true).catch(() => null),
          window.storage.get("tickets:hotels", true).catch(() => null),
          window.storage.get("tickets:visas", true).catch(() => null),
          window.storage.get("tickets:cars", true).catch(() => null),
          window.storage.get("tickets:files", true).catch(() => null),
          window.storage.get("tickets:employees", true).catch(() => null),
          window.storage.get("tickets:suggestions", true).catch(() => null),
          window.storage.get("tickets:license", true).catch(() => null),
          window.storage.get("tickets:requests", true).catch(() => null),
          window.storage.get("tickets:expenses", true).catch(() => null),
          window.storage.get("tickets:supplierPayments", true).catch(() => null),
          window.storage.get("tickets:customerPayments", true).catch(() => null),
          window.storage.get("tickets:treasuryAccounts", true).catch(() => null),
          window.storage.get("tickets:treasuryEntries", true).catch(() => null),
        ]);
        if (cancelled) return;
        if (expensesRes && expensesRes.value) {
          try { setExpenses(JSON.parse(expensesRes.value)); } catch (e) { /* ignore malformed data for this cycle */ }
        }
        if (supplierPaymentsRes && supplierPaymentsRes.value) {
          try { setSupplierPayments(JSON.parse(supplierPaymentsRes.value)); } catch (e) { /* ignore malformed data for this cycle */ }
        }
        if (customerPaymentsRes && customerPaymentsRes.value) {
          try { setCustomerPayments(JSON.parse(customerPaymentsRes.value)); } catch (e) { /* ignore malformed data for this cycle */ }
        }
        if (treasuryAccountsRes && treasuryAccountsRes.value) {
          try { setTreasuryAccounts(JSON.parse(treasuryAccountsRes.value)); } catch (e) { /* ignore malformed data for this cycle */ }
        }
        if (treasuryEntriesRes && treasuryEntriesRes.value) {
          try { setTreasuryEntries(JSON.parse(treasuryEntriesRes.value)); } catch (e) { /* ignore malformed data for this cycle */ }
        }
        if (ticketsRes && ticketsRes.value) {
          try {
            setTickets(JSON.parse(ticketsRes.value));
          } catch (e) {
            // ignore malformed data for this cycle, try again next poll
          }
        }
        if (hotelsRes && hotelsRes.value) {
          try {
            setHotelBookings(JSON.parse(hotelsRes.value));
          } catch (e) {
            // ignore malformed data for this cycle, try again next poll
          }
        }
        if (visasRes && visasRes.value) {
          try {
            setVisaBookings(JSON.parse(visasRes.value));
          } catch (e) {
            // ignore malformed data for this cycle, try again next poll
          }
        }
        if (carsRes && carsRes.value) {
          try {
            setCarBookings(JSON.parse(carsRes.value));
          } catch (e) {
            // ignore malformed data for this cycle, try again next poll
          }
        }
        if (filesRes && filesRes.value) {
          try {
            setFiles(JSON.parse(filesRes.value));
          } catch (e) {
            // ignore malformed data for this cycle, try again next poll
          }
        }
        if (employeesRes && employeesRes.value) {
          try {
            setEmployees(JSON.parse(employeesRes.value));
          } catch (e) {
            // ignore malformed data for this cycle, try again next poll
          }
        }
        if (suggestionsRes && suggestionsRes.value) {
          try {
            const parsed = JSON.parse(suggestionsRes.value);
            setSuggestions({
              companies: parsed.companies || [],
              customers: [], // never restore saved customer names — this field must have no autocomplete history
              airlines: parsed.airlines || [],
              cities: parsed.cities || [],
              suppliers: parsed.suppliers || [],
              hotelNames: parsed.hotelNames || [],
              visaSuppliers: parsed.visaSuppliers || [],
              carSuppliers: parsed.carSuppliers || [],
            });
          } catch (e) {
            // ignore malformed data for this cycle, try again next poll
          }
        }
        if (licenseRes && licenseRes.value) {
          try {
            setLicenseRecord(JSON.parse(licenseRes.value));
          } catch (e) {
            // ignore malformed data for this cycle, try again next poll
          }
        } else {
          setLicenseRecord(null);
        }
        if (requestsRes && requestsRes.value) {
          try {
            const parsedRequests = JSON.parse(requestsRes.value);
            setRequests(parsedRequests);
            const freshIncoming = parsedRequests.find(
              (r) =>
                r.toUsername === currentUser.username &&
                r.status === "pending" &&
                !seenRequestIdsRef.current.has(r.id)
            );
            parsedRequests.forEach((r) => seenRequestIdsRef.current.add(r.id));
            if (freshIncoming) {
              setIncomingRequestPopup(freshIncoming);
            }
          } catch (e) {
            // ignore malformed data for this cycle, try again next poll
          }
        }
      } catch (e) {
        // Live refresh is best-effort; a failed poll just tries again next interval
      }
    };
    const interval = setInterval(loadCoreData, LIVE_REFRESH_INTERVAL_MS);
    return () => {
      cancelled = true;
      clearInterval(interval);
    };
  }, [currentUser]);


  const ONLINE_THRESHOLD_MS = 15 * 1000; // considered "connected" if seen in the last 15s
  const HEARTBEAT_INTERVAL_MS = 5 * 1000;

  // A short, human-readable description of what this signed-in account appears to be
  // doing right now, broadcast alongside the presence heartbeat below so the main
  // account's "online now" list can show it next to each employee. Includes concrete
  // details (customer name, company, hotel, ...) rather than just the section name, so
  // the main account can tell at a glance what someone is actually working on.
  const myActivity = (() => {
    if (showManage) return "Managing employees";
    if (showManageCompanies) return "Managing companies";
    if (activeSection === "hotels") {
      if (viewingHotelBooking) {
        return `Viewing hotel booking — ${viewingHotelBooking.hotel || "hotel"}${
          viewingHotelBooking.customer ? ` (${viewingHotelBooking.customer})` : ""
        }`;
      }
      if (hotelEditingId) {
        const hb = hotelBookings.find((h) => h.id === hotelEditingId);
        return `Editing hotel booking — ${(hb && hb.hotel) || "hotel"}`;
      }
      return "Browsing hotel bookings";
    }
    if (activeSection === "visa") {
      if (visaEditingId) {
        const vb = visaBookings.find((v) => v.id === visaEditingId);
        return `Editing visa booking — ${(vb && vb.visaType) || "visa"}`;
      }
      return "Browsing visa bookings";
    }
    if (activeSection === "cars") return "Cars";
    if (activeSection === "files") return "Files";
    if (activeSection === "accounts") return "Accounts";
    // Flights (the default section)
    if (viewingTicketId) {
      const vt = tickets.find((x) => x.id === viewingTicketId);
      const vtCustomers = vt && Array.isArray(vt.customers) && vt.customers.length > 0
        ? vt.customers
        : [{ name: (vt && vt.customer) || "" }];
      const name = vtCustomers[0] && vtCustomers[0].name;
      return `Viewing ticket${name ? ` — ${name}` : ""}${vt && vt.company ? ` (${vt.company})` : ""}`;
    }
    if (form.id) {
      const name = form.customers && form.customers[0] && form.customers[0].name;
      return `Editing ticket${name ? ` — ${name}` : ""}${form.company ? ` (${form.company})` : ""}`;
    }
    // Someone with a non-empty new-ticket draft counts as actively adding one
    const draftName = form.customers && form.customers[0] && form.customers[0].name;
    if (draftName || form.company) {
      return `Adding a new ticket${draftName ? ` — ${draftName}` : ""}${form.company ? ` (${form.company})` : ""}`;
    }
    return "Browsing flights list";
  })();
  // Kept in a ref (rather than read directly) so the heartbeat interval below — which
  // only re-subscribes when currentUser changes — always sends the latest activity
  // instead of the value captured when the interval was first created.
  const myActivityRef = useRef(myActivity);
  useEffect(() => {
    myActivityRef.current = myActivity;
  }, [myActivity]);

  // While signed in, periodically mark this account as "connected" so the main account can see it
  useEffect(() => {
    if (!currentUser) return;
    let cancelled = false;
    const beat = async () => {
      try {
        await window.storage.set(
          `tickets:presence:${currentUser.username}`,
          JSON.stringify({ name: currentUser.name, ts: Date.now(), activity: myActivityRef.current }),
          true
        );
      } catch (e) {
        // Presence is a convenience feature; failures here are silent
      }
    };
    beat();
    const interval = setInterval(() => {
      if (!cancelled) beat();
    }, HEARTBEAT_INTERVAL_MS);
    return () => {
      cancelled = true;
      clearInterval(interval);
    };
  }, [currentUser]);

  // The main account (and any Owner/GM-grade employee, who gets the same admin-level
  // view) polls who else is currently connected. A plain employee record load may still
  // be in flight when this first runs, so it re-checks whenever `employees` updates too.
  useEffect(() => {
    if (!currentUser) return;
    const viewerIsOwnerGrade =
      !currentUser.isAdmin &&
      !!(employees || []).find((e) => e.username === currentUser.username && e.isOwner);
    if (!currentUser.isAdmin && !viewerIsOwnerGrade) return;
    let cancelled = false;
    const loadPresence = async () => {
      try {
        const listRes = await window.storage.list("tickets:presence:", true);
        const keys = (listRes && listRes.keys) || [];
        const entries = await Promise.all(
          keys.map(async (k) => {
            try {
              const r = await window.storage.get(k, true);
              if (!r || !r.value) return null;
              const parsed = JSON.parse(r.value);
              const username = k.replace("tickets:presence:", "");
              return [username, { ts: parsed.ts, activity: parsed.activity || "" }];
            } catch (e) {
              return null;
            }
          })
        );
        if (cancelled) return;
        const map = {};
        entries.forEach((entry) => {
          if (entry) map[entry[0]] = entry[1];
        });
        setPresenceMap(map);
      } catch (e) {
        // ignore presence load failures
      }
    };
    loadPresence();
    const interval = setInterval(loadPresence, LIVE_REFRESH_INTERVAL_MS);
    return () => {
      cancelled = true;
      clearInterval(interval);
    };
  }, [currentUser, employees]);

  const isOnline = (username) => {
    const entry = presenceMap[username];
    return !!entry && Date.now() - entry.ts < ONLINE_THRESHOLD_MS;
  };
  const onlineUsernames = Object.keys(presenceMap).filter((u) => isOnline(u));

  // Detects a remote "force sign-out": when the main account signs someone out from the
  // "online now" panel, a shared flag is written with a timestamp (see handleForceSignOut
  // below). Every signed-in client — including this one — checks its own flag on each
  // heartbeat and signs itself out automatically if the flag is newer than when this
  // particular session started (so it can never retroactively kill a brand-new login).
  useEffect(() => {
    if (!currentUser) return;
    let cancelled = false;
    const checkForceLogout = async () => {
      try {
        const res = await window.storage.get(`tickets:forceLogout:${currentUser.username}`, true).catch(() => null);
        if (cancelled || !res || !res.value) return;
        const ts = parseInt(res.value, 10);
        if (ts && ts > sessionStartedAtRef.current) {
          await handleLogout();
        }
      } catch (e) {
        // Best-effort; a missed check just retries on the next heartbeat
      }
    };
    const interval = setInterval(() => {
      if (!cancelled) checkForceLogout();
    }, HEARTBEAT_INTERVAL_MS);
    return () => {
      cancelled = true;
      clearInterval(interval);
    };
  }, [currentUser]);

  // Clears the "online now" presence flag when the page/tab is closed or navigated away
  // from, so this employee stops showing as online right away. Deliberately does NOT
  // touch the saved session here — "beforeunload"/"pagehide" also fire on a normal page
  // refresh, and clearing the session there was signing people out just from reloading
  // the page. Signing out now only happens via the explicit Sign out button, a remote
  // force-sign-out, or the inactivity timeout below.
  useEffect(() => {
    if (!currentUser) return;
    const username = currentUser.username;
    const handleUnload = () => {
      try { window.storage.delete(`tickets:presence:${username}`, true); } catch (e) {}
    };
    window.addEventListener("pagehide", handleUnload);
    window.addEventListener("beforeunload", handleUnload);
    return () => {
      window.removeEventListener("pagehide", handleUnload);
      window.removeEventListener("beforeunload", handleUnload);
    };
  }, [currentUser]);

  // Auto sign-out after 30 minutes of inactivity. Any mouse, keyboard, scroll, or touch
  // activity resets the timer; if it ever fires, the session is ended the same way the
  // Sign out button does it.
  const INACTIVITY_TIMEOUT_MS = 30 * 60 * 1000;
  useEffect(() => {
    if (!currentUser) return;
    let timeoutId;
    const resetTimer = () => {
      clearTimeout(timeoutId);
      timeoutId = setTimeout(() => {
        handleLogout();
      }, INACTIVITY_TIMEOUT_MS);
    };
    const activityEvents = ["mousedown", "mousemove", "keydown", "wheel", "touchstart", "scroll"];
    activityEvents.forEach((evt) => window.addEventListener(evt, resetTimer, { passive: true }));
    resetTimer();
    return () => {
      clearTimeout(timeoutId);
      activityEvents.forEach((evt) => window.removeEventListener(evt, resetTimer));
    };
  }, [currentUser]);

  const persistTickets = async (next) => {
    setTickets(next);
    try {
      await window.storage.set("tickets:list", JSON.stringify(next), true);
    } catch (e) {
      setError("Could not save data, please try again");
    }
  };

  // The USD -> EGP rate is entered by hand (e.g. from the CBE's published rate each
  // morning) and saved to shared storage, so every signed-in employee sees the same
  // rate without each of them having to type it in separately.
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("tickets:usdRate", true).catch(() => null);
        if (res && res.value) {
          const parsed = JSON.parse(res.value);
          setUsdToEgpRate(parsed.rate ?? null);
          setUsdToEgpRateDate(parsed.date || "");
        }
      } catch (e) {
        // no saved rate yet
      }
    })();
  }, []);

  const persistUsdRate = async (rate) => {
    const date = todayDateStr();
    setUsdToEgpRate(rate);
    setUsdToEgpRateDate(date);
    try {
      await window.storage.set("tickets:usdRate", JSON.stringify({ rate, date }), true);
    } catch (e) {
      // Saving the rate is best-effort; the typed value still applies locally either way
    }
  };

  // IATA balance — same shared-storage pattern as the USD rate above, so every signed-in
  // employee sees the same running balance. iataBalanceLoaded is set true once this fetch
  // settles (found a value or not) — deductions are blocked until then, see
  // applyIataTicketValue below.
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("tickets:iataBalance", true).catch(() => null);
        if (res && res.value !== undefined && res.value !== null && res.value !== "") {
          const parsed = parseFloat(res.value);
          if (!Number.isNaN(parsed)) setIataBalance(parsed);
        }
      } catch (e) {
        // no saved balance yet
      } finally {
        setIataBalanceLoaded(true);
      }
    })();
  }, []);

  // Loads the IATA history log the same way the balance itself is loaded above. The log
  // only ever holds today's deductions — if the saved entry is from an earlier day, it's
  // treated as empty rather than shown, since history starts fresh each day.
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("tickets:iataHistory", true).catch(() => null);
        if (res && res.value) {
          const parsed = JSON.parse(res.value);
          if (parsed && parsed.date === todayDateStr() && Array.isArray(parsed.deductions)) {
            setIataHistory(parsed);
          }
        }
      } catch (e) {
        // no saved history yet
      }
    })();
  }, []);

  const persistIataHistory = async (next) => {
    setIataHistory(next);
    try {
      await window.storage.set("tickets:iataHistory", JSON.stringify(next), true);
    } catch (e) {
      // Saving is best-effort; the list still applies locally either way
    }
  };

  // Logs one deducted amount under today's date. The log holds only today's deductions —
  // if the last saved entry is from a previous day, it's dropped and today starts empty,
  // so the History popup never carries anything over from an earlier day. Only actual
  // deductions are logged here — manually overwriting the balance box itself is not.
  const recordIataDeduction = (amount, balanceBefore, balanceAfter) => {
    const today = todayDateStr();
    const entry = { amount, balanceBefore, balanceAfter, time: new Date().toISOString() };
    const sameDay = iataHistory && iataHistory.date === today && Array.isArray(iataHistory.deductions);
    const next = sameDay
      ? { date: today, deductions: [...iataHistory.deductions, entry] }
      : { date: today, deductions: [entry] };
    persistIataHistory(next);
  };

  const persistIataBalance = async (balance) => {
    setIataBalance(balance);
    try {
      await window.storage.set("tickets:iataBalance", String(balance), true);
    } catch (e) {
      // Saving is best-effort; the value still applies locally either way
    }
  };

  // Deducts the amount typed into the "Issued ticket value" box from the IATA balance,
  // logs it in today's history, then clears the box so it's ready for the next ticket.
  // Pressing Enter in that box (see onKeyDown below) is the only way this fires —
  // there's no separate Deduct button. Blocked until the saved balance has actually
  // loaded, so it can never compute against an unloaded `null` and wipe out the real
  // balance — the balance itself only ever changes here or when typed directly by hand.
  const applyIataTicketValue = () => {
    if (!iataBalanceLoaded) return;
    const val = parseFloat(iataTicketValueInput);
    if (Number.isNaN(val) || val === 0) return;
    const balanceBefore = iataBalance || 0;
    const nextBalance = balanceBefore - val;
    persistIataBalance(nextBalance);
    recordIataDeduction(val, balanceBefore, nextBalance);
    setIataTicketValueInput("");
  };

  const persistHotelBookings = async (next) => {
    setHotelBookings(next);
    try {
      await window.storage.set("tickets:hotels", JSON.stringify(next), true);
    } catch (e) {
      setHotelError("Could not save data, please try again");
    }
  };

  const persistVisaBookings = async (next) => {
    setVisaBookings(next);
    try {
      await window.storage.set("tickets:visas", JSON.stringify(next), true);
    } catch (e) {
      setVisaError("Could not save data, please try again");
    }
  };

  const persistCarBookings = async (next) => {
    setCarBookings(next);
    try {
      await window.storage.set("tickets:cars", JSON.stringify(next), true);
    } catch (e) {
      setCarError("Could not save data, please try again");
    }
  };

  const persistFiles = async (next) => {
    setFiles(next);
    try {
      await window.storage.set("tickets:files", JSON.stringify(next), true);
    } catch (e) {
      setFileError("Could not save data, please try again");
    }
  };

  const persistEmployees = async (next) => {
    setEmployees(next);
    try {
      await window.storage.set("tickets:employees", JSON.stringify(next), true);
    } catch (e) {
      setManageError("Could not save the employee list, please try again");
    }
  };

  const persistSuggestions = async (next) => {
    setSuggestions(next);
    try {
      await window.storage.set("tickets:suggestions", JSON.stringify(next), true);
    } catch (e) {
      // Suggestions are a convenience feature, so failures here are silent
    }
  };

  const persistRequests = async (next) => {
    setRequests(next);
    next.forEach((r) => seenRequestIdsRef.current.add(r.id));
    try {
      await window.storage.set("tickets:requests", JSON.stringify(next), true);
    } catch (e) {
      setRequestSendError("Could not save the request, please try again");
    }
  };

  const persistExpenses = async (next) => {
    setExpenses(next);
    try {
      await window.storage.set("tickets:expenses", JSON.stringify(next), true);
    } catch (e) {
      setAccountsError("Could not save data, please try again");
    }
  };
  const persistSupplierPayments = async (next) => {
    setSupplierPayments(next);
    try {
      await window.storage.set("tickets:supplierPayments", JSON.stringify(next), true);
    } catch (e) {
      setAccountsError("Could not save data, please try again");
    }
  };
  const persistCustomerPayments = async (next) => {
    setCustomerPayments(next);
    try {
      await window.storage.set("tickets:customerPayments", JSON.stringify(next), true);
    } catch (e) {
      setAccountsError("Could not save data, please try again");
    }
  };
  const persistTreasuryAccounts = async (next) => {
    setTreasuryAccounts(next);
    try {
      await window.storage.set("tickets:treasuryAccounts", JSON.stringify(next), true);
    } catch (e) {
      setAccountsError("Could not save data, please try again");
    }
  };
  const persistTreasuryEntries = async (next) => {
    setTreasuryEntries(next);
    try {
      await window.storage.set("tickets:treasuryEntries", JSON.stringify(next), true);
    } catch (e) {
      setAccountsError("Could not save data, please try again");
    }
  };

  // Sends a new request from the signed-in account to another employee.
  const handleSendRequest = async () => {
    setRequestSendError("");
    if (!newRequestTo) {
      setRequestSendError("Please choose who this request is for");
      return;
    }
    if (!newRequestMessage.trim()) {
      setRequestSendError("Please write what you need");
      return;
    }
    const target = (employees || []).find((e) => e.username === newRequestTo);
    if (!target) {
      setRequestSendError("That employee could not be found");
      return;
    }
    const newReq = {
      id: `req_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
      fromUsername: currentUser.username,
      fromName: currentUser.name,
      toUsername: target.username,
      toName: target.name,
      message: newRequestMessage.trim(),
      status: "pending", // "pending" | "completed" | "declined"
      createdAt: new Date().toISOString(),
      respondedAt: null,
    };
    await persistRequests([newReq, ...(requests || [])]);
    setNewRequestTo("");
    setNewRequestMessage("");
  };

  // Lets the recipient of a request mark it completed or declined. Also used to
  // dismiss the incoming popup without changing its status (status stays "pending").
  const handleRespondToRequest = async (requestId, status) => {
    const next = (requests || []).map((r) =>
      r.id === requestId ? { ...r, status, respondedAt: new Date().toISOString() } : r
    );
    await persistRequests(next);
  };

  // Remembers values entered on a ticket (airline, cities) so they keep showing up as
  // autocomplete options later, even if this ticket gets deleted. Companies are
  // intentionally excluded — a new company can only be registered via the
  // "Manage companies" button, never auto-added just by typing a new name on a ticket.
  const rememberSuggestionsFromRecord = (record) => {
    const addUnique = (list, value) => {
      const v = (value || "").trim();
      if (!v) return list;
      return list.some((existing) => existing.toLowerCase() === v.toLowerCase()) ? list : [...list, v];
    };
    let next = {
      companies: [...suggestions.companies],
      // Customer names are intentionally never remembered here — the customer field
      // must never offer autocomplete/history of previously typed names.
      customers: [],
      airlines: [...suggestions.airlines],
      cities: [...suggestions.cities],
      suppliers: [...(suggestions.suppliers || [])],
      hotelNames: [...(suggestions.hotelNames || [])],
    };
    next.airlines = addUnique(next.airlines, record.airline);
    next.cities = addUnique(next.cities, record.from);
    next.cities = addUnique(next.cities, record.to);
    if (Array.isArray(record.destinations)) {
      record.destinations.forEach((d) => { next.cities = addUnique(next.cities, d); });
    }
    persistSuggestions(next);
  };


  // Lets an admin (or an employee granted the Manage companies permission) register a
  // company's full details — name, tax number, commercial registration number, and phone
  // numbers — so they're always available to pick from the Company field and filter, even
  // before any ticket has been entered for them. If editingCompanyName is set, this saves
  // changes to that existing record instead of adding a new one.
  const handleAddCompany = () => {
    if (!canManageCompanies) return;
    const name = newCompanyDraft.name.trim();
    if (!name) return;
    const duplicate = suggestions.companies.some(
      (c) =>
        companyName(c).toLowerCase() === name.toLowerCase() &&
        companyName(c).toLowerCase() !== (editingCompanyName || "").toLowerCase()
    );
    if (duplicate) {
      setCompanyError("A company with that name already exists");
      return;
    }
    const record = {
      name,
      taxNumber: newCompanyDraft.taxNumber.trim(),
      commercialReg: newCompanyDraft.commercialReg.trim(),
      phones: newCompanyDraft.phones
        .split(/[,\n]/)
        .map((p) => p.trim())
        .filter(Boolean),
    };
    const companies = editingCompanyName
      ? suggestions.companies.map((c) => (companyName(c) === editingCompanyName ? record : c))
      : [...suggestions.companies, record];
    persistSuggestions({ ...suggestions, companies });
    setNewCompanyDraft(emptyCompanyDraft);
    setEditingCompanyName(null);
    setCompanyError("");
  };

  // Loads an existing company's saved details back into the form so they can be edited.
  const handleEditCompanyClick = (c) => {
    setEditingCompanyName(companyName(c));
    setNewCompanyDraft({
      name: companyName(c),
      taxNumber: typeof c === "object" ? c.taxNumber || "" : "",
      commercialReg: typeof c === "object" ? c.commercialReg || "" : "",
      phones: typeof c === "object" && Array.isArray(c.phones) ? c.phones.join(", ") : "",
    });
  };

  const cancelEditCompany = () => {
    setEditingCompanyName(null);
    setNewCompanyDraft(emptyCompanyDraft);
  };

  // Removes a company from the saved suggestions list. Existing tickets already
  // recorded under that company name are untouched — this only affects the picker.
  const handleDeleteCompany = (name) => {
    if (!canManageCompanies) return;
    persistSuggestions({
      ...suggestions,
      companies: suggestions.companies.filter((c) => companyName(c) !== name),
    });
    if (editingCompanyName === name) cancelEditCompany();
  };

  const profit = (net, sold) => {
    const n = parseFloat(net) || 0;
    const s = parseFloat(sold) || 0;
    return s - n;
  };

  // ---------- Hotels ----------
  const resetHotelForm = () => {
    setHotelForm(getEmptyHotelForm());
    setHotelEditingId(null);
    setHotelError("");
    setHotelSupplierOther(false);
    setHotelNameOther(false);
  };

  const addHotelRoomLine = () => {
    setHotelForm({ ...hotelForm, roomLines: [...hotelForm.roomLines, emptyRoomLine()] });
  };

  const removeHotelRoomLine = (lineId) => {
    if (hotelForm.roomLines.length <= 1) return; // always keep at least one line
    setHotelForm({ ...hotelForm, roomLines: hotelForm.roomLines.filter((l) => l.id !== lineId) });
  };

  const updateHotelRoomLine = (lineId, patch) => {
    setHotelForm({
      ...hotelForm,
      roomLines: hotelForm.roomLines.map((l) => (l.id === lineId ? { ...l, ...patch } : l)),
    });
  };

  // Updates one adult guest's name within a room line, by that guest's position.
  const updateRoomGuest = (lineId, guestIndex, name) => {
    const line = hotelForm.roomLines.find((l) => l.id === lineId);
    if (!line) return;
    const guests = (line.guests || []).map((g, i) => (i === guestIndex ? { ...g, name } : g));
    updateHotelRoomLine(lineId, { guests });
  };

  const addRoomChild = (lineId) => {
    const line = hotelForm.roomLines.find((l) => l.id === lineId);
    if (!line) return;
    updateHotelRoomLine(lineId, { children: [...(line.children || []), emptyChild()] });
  };

  const updateRoomChild = (lineId, childId, patch) => {
    const line = hotelForm.roomLines.find((l) => l.id === lineId);
    if (!line) return;
    updateHotelRoomLine(lineId, {
      children: (line.children || []).map((c) => (c.id === childId ? { ...c, ...patch } : c)),
    });
  };

  const removeRoomChild = (lineId, childId) => {
    const line = hotelForm.roomLines.find((l) => l.id === lineId);
    if (!line) return;
    updateHotelRoomLine(lineId, { children: (line.children || []).filter((c) => c.id !== childId) });
  };

  const handleSaveHotel = async () => {
    setHotelError("");
    // Company name and supplier are both optional — a blank company means an Individual
    // booking, and a blank supplier just means none was specified, so neither is part of
    // the required-fields check below.
    if (!hotelForm.hotel.trim()) {
      setHotelError("Please fill in the hotel field");
      return;
    }
    const lines = hotelForm.roomLines || [];
    if (lines.length === 0) {
      setHotelError("Please add at least one room line");
      return;
    }
    for (const l of lines) {
      if ((parseInt(l.count, 10) || 0) < 1) {
        setHotelError("Each room line needs at least 1 room");
        return;
      }
      if (l.netPrice === "" || l.soldPrice === "") {
        setHotelError("Please fill in the net and sold price for every room line");
        return;
      }
      if (!l.checkIn || !l.checkOut) {
        setHotelError("Please fill in the check-in and check-out dates for every room");
        return;
      }
      if (new Date(l.checkOut) < new Date(l.checkIn)) {
        setHotelError("Check-out date can't be before check-in date for a room");
        return;
      }
      // Only the first guest in each room is required — the rest are optional.
      if (!l.guests || !l.guests[0] || !l.guests[0].name.trim()) {
        setHotelError("Please enter at least the first guest's name for every room");
        return;
      }
    }

    if (hotelEditingId) {
      const next = hotelBookings.map((h) =>
        h.id === hotelEditingId ? { ...h, ...hotelForm, id: hotelEditingId } : h
      );
      await persistHotelBookings(next);
    } else {
      const record = {
        ...hotelForm,
        id: `H-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        employee: currentUser.name,
        employeeUsername: currentUser.username,
      };
      await persistHotelBookings([record, ...hotelBookings]);
    }
    resetHotelForm();
  };

  const handleEditHotelClick = (h) => {
    setHotelEditingId(h.id);
    setHotelForm({
      id: h.id,
      employee: h.employee || "",
      customer: h.customer || "",
      hotel: h.hotel || "",
      supplier: h.supplier || "",
      roomLines:
        Array.isArray(h.roomLines) && h.roomLines.length > 0
          ? h.roomLines.map((l) => ({
              ...l,
              id: l.id || emptyRoomLine().id,
              currency: l.currency || h.currency || "EGP",
              // Legacy bookings kept dates on the booking itself rather than per room —
              // fall back to those so older records still show something sensible.
              checkIn: l.checkIn || h.checkIn || todayDateStr(),
              checkOut: l.checkOut || h.checkOut || todayDateStr(),
              // Legacy bookings had no guest names — pad an empty list to match capacity.
              guests: guestsForCapacity(l.guests, ROOM_CAPACITY[l.roomType] || 1),
              children: Array.isArray(l.children) ? l.children : [],
            }))
          : [emptyRoomLine()],
      bookingDate: h.bookingDate || todayDateStr(),
      notes: h.notes || "",
    });
    setHotelSupplierOther(!!h.supplier && !suggestions.suppliers.includes(h.supplier));
    setHotelNameOther(!!h.hotel && !suggestions.hotelNames.includes(h.hotel));
    setHotelError("");
  };

  const handleDeleteHotel = (id, onDeleted) => {
    requestConfirm("Delete this hotel booking? This cannot be undone.", async () => {
      await persistHotelBookings(hotelBookings.filter((h) => h.id !== id));
      if (hotelEditingId === id) resetHotelForm();
      setConfirmDialog(null);
      if (onDeleted) onDeleted();
    });
  };

  // Registers a new supplier name so it's always available to pick from the Hotels
  // page's Supplier field, via the "+ Add supplier" button at the top of the page.
  const handleAddSupplierName = () => {
    const name = newSupplierDraft.trim();
    if (!name) return;
    const duplicate = (suggestions.suppliers || []).some((s) => s.toLowerCase() === name.toLowerCase());
    if (duplicate) {
      setHotelError("This supplier already exists");
      return;
    }
    persistSuggestions({ ...suggestions, suppliers: [...(suggestions.suppliers || []), name] });
    setNewSupplierDraft("");
    setHotelError("");
  };

  const handleDeleteSupplierName = (name) => {
    persistSuggestions({ ...suggestions, suppliers: (suggestions.suppliers || []).filter((s) => s !== name) });
  };

  // Registers a new hotel name so it's always available to pick from the Hotels
  // page's Hotel name field, via the "+ Add hotel name" button at the top of the page.
  const handleAddHotelName = () => {
    const name = newHotelNameDraft.trim();
    if (!name) return;
    const duplicate = (suggestions.hotelNames || []).some((h) => h.toLowerCase() === name.toLowerCase());
    if (duplicate) {
      setHotelError("This hotel already exists");
      return;
    }
    persistSuggestions({ ...suggestions, hotelNames: [...(suggestions.hotelNames || []), name] });
    setNewHotelNameDraft("");
    setHotelError("");
  };

  const handleDeleteHotelName = (name) => {
    persistSuggestions({ ...suggestions, hotelNames: (suggestions.hotelNames || []).filter((h) => h !== name) });
  };

  // ---------- Visa ----------
  const resetVisaForm = () => {
    setVisaForm(getEmptyVisaForm());
    setVisaEditingId(null);
    setVisaError("");
    setVisaSupplierOther(false);
  };

  const handleVisaCustomersCountChange = (value) => {
    const count = parseInt(value, 10) || 1;
    const customers = resizeVisaCustomers(visaForm.customers, count);
    setVisaForm({ ...visaForm, customersCount: count, customers });
  };

  const handleVisaCustomerNameChange = (index, name) => {
    const customers = visaForm.customers.map((c, i) => (i === index ? { ...c, name } : c));
    setVisaForm({ ...visaForm, customers });
  };

  const handleSaveVisa = async () => {
    setVisaError("");
    if (!visaForm.visaType.trim()) {
      setVisaError("Please fill in the visa field");
      return;
    }
    if (!visaForm.customers[0] || !visaForm.customers[0].name.trim()) {
      setVisaError("Please enter at least the first customer's name");
      return;
    }
    if (visaForm.netPrice === "" || visaForm.soldPrice === "") {
      setVisaError("Please fill in the net and sold prices");
      return;
    }
    if (visaEditingId) {
      const next = visaBookings.map((v) => (v.id === visaEditingId ? { ...v, ...visaForm, id: visaEditingId } : v));
      await persistVisaBookings(next);
    } else {
      const record = {
        ...visaForm,
        id: `V-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        employee: currentUser.name,
        employeeUsername: currentUser.username,
      };
      await persistVisaBookings([record, ...visaBookings]);
    }
    resetVisaForm();
  };

  const handleEditVisaClick = (v) => {
    setVisaEditingId(v.id);
    setVisaForm({
      id: v.id,
      customersCount: (v.customers || []).length || 1,
      customers: v.customers && v.customers.length > 0 ? v.customers.map((c) => ({ ...c })) : [emptyVisaCustomer()],
      visaType: v.visaType || "",
      supplier: v.supplier || "",
      currency: v.currency || "EGP",
      netPrice: v.netPrice,
      soldPrice: v.soldPrice,
      bookingDate: v.bookingDate || todayDateStr(),
    });
    setVisaSupplierOther(!!v.supplier && !(suggestions.visaSuppliers || []).includes(v.supplier));
    setVisaError("");
  };

  const handleDeleteVisa = (id, onDeleted) => {
    requestConfirm("Delete this visa booking? This cannot be undone.", async () => {
      await persistVisaBookings(visaBookings.filter((v) => v.id !== id));
      if (visaEditingId === id) resetVisaForm();
      setConfirmDialog(null);
      if (onDeleted) onDeleted();
    });
  };

  // Registers a new supplier name in the Visa page's OWN supplier list — kept separate
  // from the Hotels/Flights supplier lists, via the "+ Add supplier" button at the top
  // of the Visa page.
  const handleAddVisaSupplierName = () => {
    const name = newVisaSupplierDraft.trim();
    if (!name) return;
    const duplicate = (suggestions.visaSuppliers || []).some((s) => s.toLowerCase() === name.toLowerCase());
    if (duplicate) {
      setVisaError("This supplier already exists");
      return;
    }
    persistSuggestions({ ...suggestions, visaSuppliers: [...(suggestions.visaSuppliers || []), name] });
    setNewVisaSupplierDraft("");
    setVisaError("");
  };

  const handleDeleteVisaSupplierName = (name) => {
    persistSuggestions({ ...suggestions, visaSuppliers: (suggestions.visaSuppliers || []).filter((s) => s !== name) });
  };

  const resetCarForm = () => {
    setCarForm(getEmptyCarForm());
    setCarEditingId(null);
    setCarError("");
    setCarSupplierOther(false);
  };

  const handleSaveCar = async () => {
    setCarError("");
    if (!carForm.customerName.trim()) {
      setCarError("Please enter the customer name");
      return;
    }
    if (!carForm.routeFrom.trim() || !carForm.routeTo.trim()) {
      setCarError("Please fill in the route (from and to)");
      return;
    }
    if (!carForm.carType.trim()) {
      setCarError("Please fill in the car type");
      return;
    }
    if (carForm.startsAtAirport && !carForm.flightNumber.trim()) {
      setCarError("Please enter the flight number");
      return;
    }
    if (carForm.netPrice === "" || carForm.soldPrice === "") {
      setCarError("Please fill in the net and sold prices");
      return;
    }
    if (carEditingId) {
      const next = carBookings.map((c) => (c.id === carEditingId ? { ...c, ...carForm, id: carEditingId } : c));
      await persistCarBookings(next);
    } else {
      const record = {
        ...carForm,
        id: `C-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        employee: currentUser.name,
        employeeUsername: currentUser.username,
      };
      await persistCarBookings([record, ...carBookings]);
    }
    resetCarForm();
  };

  const handleEditCarClick = (c) => {
    setCarEditingId(c.id);
    setCarForm({
      id: c.id,
      customerName: c.customerName || "",
      phone: c.phone || "",
      routeFrom: c.routeFrom || "",
      routeTo: c.routeTo || "",
      carType: c.carType || "",
      supplier: c.supplier || "",
      hasWaiting: !!c.hasWaiting,
      waitingHours: c.waitingHours || "",
      isRoundTrip: !!c.isRoundTrip,
      driverTip: c.driverTip || "",
      startsAtAirport: !!c.startsAtAirport,
      flightNumber: c.flightNumber || "",
      currency: c.currency || "EGP",
      netPrice: c.netPrice,
      soldPrice: c.soldPrice,
      bookingDate: c.bookingDate || todayDateStr(),
      bookingTime: c.bookingTime || "",
      returnDate: c.returnDate || "",
      returnTime: c.returnTime || "",
      entryDate: c.entryDate || todayDateStr(),
      collection: c.collection || "",
    });
    setCarSupplierOther(!!c.supplier && !(suggestions.carSuppliers || []).includes(c.supplier));
    setCarError("");
  };

  const handleDeleteCar = (id, onDeleted) => {
    requestConfirm("Delete this transfer booking? This cannot be undone.", async () => {
      await persistCarBookings(carBookings.filter((c) => c.id !== id));
      if (carEditingId === id) resetCarForm();
      setConfirmDialog(null);
      if (onDeleted) onDeleted();
    });
  };

  // Shared layout for every printable receipt (transfers, hotels, visa). `sections` is
  // an array of { heading, rows: [[label, value], ...] } — keeping this in one place
  // means every service's receipt looks and behaves the same.
  const buildReceiptHtml = (docTitle, subtitle, sections) => {
    const printedBy = currentUser?.name || "";

    const row = (label, value) =>
      value === "" || value === null || value === undefined
        ? ""
        : `<tr><td class="label">${label}</td><td class="value">${value}</td></tr>`;

    const sectionsHtml = sections
      .map(
        (s) => `
          <h2>${s.heading}</h2>
          <table>${s.rows.map(([label, value]) => row(label, value)).join("")}</table>`
      )
      .join("");

    return `
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="utf-8" />
          <title>${docTitle}</title>
          <style>
            * { box-sizing: border-box; }
            body { font-family: Arial, Helvetica, sans-serif; color: #292524; padding: 32px; }
            .header { display: flex; align-items: center; gap: 16px; border-bottom: 2px solid #115e59; padding-bottom: 16px; margin-bottom: 24px; }
            .header img { width: 90px; height: auto; object-fit: contain; }
            .header h1 { font-size: 20px; margin: 0; color: #115e59; }
            .header p { margin: 2px 0 0; font-size: 12px; color: #78716c; }
            h2 { font-size: 14px; color: #115e59; margin: 20px 0 8px; text-transform: uppercase; letter-spacing: 0.05em; }
            table { width: 100%; border-collapse: collapse; font-size: 13px; }
            td.label { padding: 6px 10px; color: #78716c; width: 40%; border-bottom: 1px solid #e7e5e4; }
            td.value { padding: 6px 10px; font-weight: 600; border-bottom: 1px solid #e7e5e4; }
            .footer { margin-top: 32px; font-size: 11px; color: #a8a29e; text-align: right; }
            @media print {
              body { padding: 0 24px; }
            }
          </style>
        </head>
        <body>
          <div class="header">
            <img src="${LOGO_DATA_URL}" alt="Perla Di Mare" />
            <div>
              <h1>${subtitle}</h1>
              <p>Perla Di Mare</p>
            </div>
          </div>
          ${sectionsHtml}
          <div class="footer">
            ${printedBy ? `Printed by ${printedBy} &middot; ` : ""}${new Date().toLocaleString()}
          </div>
        </body>
      </html>
    `;
  };

  // Opens the print preview popup (see printPreview state above) instead of a separate
  // browser tab/window — this is the in-app "popup system" used for every printable
  // receipt across all services.
  const openPrintPreview = (docTitle, subtitle, sections) => {
    setPrintPreview({ title: docTitle, html: buildReceiptHtml(docTitle, subtitle, sections) });
  };

  // Opens a printable receipt for a single flight ticket (with its refund, if any) in
  // the print preview popup.
  const handlePrintTicket = (t) => {
    const customers = getCustomers(t);
    const customerRows = customers.map((c, i) => [
      `Customer ${i + 1}`,
      `${c.name || "-"}${c.ticketNumber ? ` — ${c.ticketNumber}` : ""}${c.pnrReference ? ` (PNR: ${c.pnrReference})` : ""}`,
    ]);

    const sections = [
      {
        heading: "Booking",
        rows: [
          ["Company", t.company && t.company.trim() ? t.company : "Individual"],
          ["Supplier", t.supplier || "-"],
          ["Route", routeLabel(t)],
          ["Airline", t.airline ? (getAirlineIata(t.airline) || t.airline) : "-"],
          ["Ticket issue date", t.date ? formatDisplayDate(t.date) : "-"],
          ...(t.isReissued ? [["Exchanged from", t.oldTicketNumber || "an older ticket"]] : []),
        ],
      },
      {
        heading: "Customers",
        rows: customerRows,
      },
      {
        heading: "Pricing",
        rows: [
          ["Net price", fmt(t.netPrice)],
          ["Sold price", fmt(t.soldPrice)],
          ["Profit", fmt(profit(t.netPrice, t.soldPrice))],
        ],
      },
      ...(hasRefund(t)
        ? [
            {
              heading: "Refund",
              rows: [
                ["Refunded by airline", fmt(getRefunds(t)[0]?.airlineAmount)],
                ["Refunded to customer", fmt(getRefunds(t)[0]?.customerAmount)],
                ["Net after refund", fmt(netAfterRefund(t))],
                ["Sold after refund", fmt(soldAfterRefund(t))],
                ["Profit after refund", fmt(profitAfterRefund(t))],
              ],
            },
          ]
        : []),
      ...(t.notes ? [{ heading: "Notes", rows: [["Notes", t.notes]] }] : []),
    ];

    openPrintPreview(`Ticket - ${(customers[0] && customers[0].name) || ""}`, "Flight Ticket Receipt", sections);
  };

  // Opens a printable receipt for a single transfer booking in the print preview popup.
  const handlePrintCar = (c) => {
    openPrintPreview(`Transfer Booking - ${c.customerName || ""}`, "Transfer Booking Receipt", [
      {
        heading: "Customer",
        rows: [
          ["Customer name", c.customerName || "-"],
          ["Phone", c.phone || "-"],
        ],
      },
      {
        heading: "Transfer details",
        rows: [
          ["Route", `${c.routeFrom || "-"} \u2192 ${c.routeTo || "-"}`],
          ["Car type", c.carType || "-"],
          ["Supplier", c.supplier || "-"],
          ["Trip", c.isRoundTrip ? "Round trip" : "One way"],
          ["Waiting", c.hasWaiting ? `${c.waitingHours || 0} h` : "-"],
          ["Flight number", c.startsAtAirport ? (c.flightNumber || "-") : "-"],
          ["Booking date", c.bookingDate ? formatDisplayDate(c.bookingDate) : "-"],
          ["Booking time", c.bookingTime || "-"],
          ...(c.isRoundTrip
            ? [
                ["Return date", c.returnDate ? formatDisplayDate(c.returnDate) : "-"],
                ["Return time", c.returnTime || "-"],
              ]
            : []),
          ["Collection", c.collection ? `${fmt(parseFloat(c.collection) || 0)} ${c.currency}` : "-"],
          ["Driver tip", c.driverTip ? `${fmt(parseFloat(c.driverTip) || 0)} ${c.currency}` : "-"],
        ],
      },
    ]);
  };

  // Opens a printable receipt for a single hotel booking (all its room lines) in the
  // print preview popup.
  const handlePrintHotel = (h) => {
    const roomSections = (h.roomLines || []).map((line, i) => {
      const roomLabel = (ROOM_TYPES.find((r) => r.value === line.roomType) || {}).label || line.roomType || "-";
      const mealLabel = (MEAL_PLANS.find((m) => m.value === line.mealPlan) || {}).label || line.mealPlan || "-";
      const nights = roomLineNights(line, h);
      const guestNames = (line.guests || []).map((g) => g.name).filter(Boolean).join(", ") || "-";
      const childrenText =
        (line.children || [])
          .filter((ch) => ch.name)
          .map((ch) => `${ch.name} (${ch.age !== "" && ch.age != null ? ch.age : "-"}y)`)
          .join(", ") || "-";
      return {
        heading: `Room ${i + 1} \u2014 ${line.count || 1}x ${roomLabel}`,
        rows: [
          ["Meal plan", mealLabel],
          ["Check-in", line.checkIn ? formatDisplayDate(line.checkIn) : "-"],
          ["Check-out", line.checkOut ? formatDisplayDate(line.checkOut) : "-"],
          ["Nights", nights],
          ["Guests", guestNames],
          ["Children", childrenText],
          ["Net (per room/night)", `${fmt(hotelLineNetTotal(line, nights))} ${line.currency}`],
          ["Sold (per room/night)", `${fmt(hotelLineSoldTotal(line, nights))} ${line.currency}`],
        ],
      };
    });

    openPrintPreview(`Hotel Booking - ${h.hotel || ""}`, "Hotel Booking Receipt", [
      {
        heading: "Booking",
        rows: [
          ["Company", h.customer && h.customer.trim() ? h.customer : "Individual"],
          ["Hotel", h.hotel || "-"],
          ["Supplier", h.supplier || "-"],
          ["Booking date", h.bookingDate ? formatDisplayDate(h.bookingDate) : "-"],
          ["Notes", h.notes || "-"],
        ],
      },
      ...roomSections,
      {
        heading: "Totals",
        rows: [
          ["Net total", `${fmt(hotelNetTotal(h))} EGP`],
          ["Sold total", `${fmt(hotelSoldTotal(h))} EGP`],
          ["Profit", `${fmt(hotelProfitTotal(h))} EGP`],
        ],
      },
    ]);
  };

  // Opens a printable receipt for a single visa booking in the print preview popup.
  const handlePrintVisa = (v) => {
    const customerNames = (v.customers || []).map((c) => c.name || "-").join(", ") || "-";
    openPrintPreview(
      `Visa Booking - ${(v.customers && v.customers[0] && v.customers[0].name) || ""}`,
      "Visa Booking Receipt",
      [
        {
          heading: "Visa details",
          rows: [
            ["Visa", v.visaType || "-"],
            ["Supplier", v.supplier || "-"],
            ["Booking date", v.bookingDate ? formatDisplayDate(v.bookingDate) : "-"],
            ["Number of customers", (v.customers || []).length || 1],
            ["Customers", customerNames],
          ],
        },
        {
          heading: "Pricing",
          rows: [
            ["Net (per person)", `${fmt(parseFloat(v.netPrice) || 0)} ${v.currency}`],
            ["Sold (per person)", `${fmt(parseFloat(v.soldPrice) || 0)} ${v.currency}`],
            ["Net total", `${fmt(visaNetTotal(v))} ${v.currency}`],
            ["Sold total", `${fmt(visaSoldTotal(v))} ${v.currency}`],
            ["Profit", `${fmt(visaProfitTotal(v))} ${v.currency}`],
          ],
        },
      ]
    );
  };

  // Registers a new supplier name in the Transfers page's OWN supplier list — kept
  // separate from the Hotels/Flights/Visa supplier lists, via the "+ Add supplier"
  // button at the top of the Transfers page.
  const handleAddCarSupplierName = () => {
    const name = newCarSupplierDraft.trim();
    if (!name) return;
    const duplicate = (suggestions.carSuppliers || []).some((s) => s.toLowerCase() === name.toLowerCase());
    if (duplicate) {
      setCarError("This supplier already exists");
      return;
    }
    persistSuggestions({ ...suggestions, carSuppliers: [...(suggestions.carSuppliers || []), name] });
    setNewCarSupplierDraft("");
    setCarError("");
  };

  const handleDeleteCarSupplierName = (name) => {
    persistSuggestions({ ...suggestions, carSuppliers: (suggestions.carSuppliers || []).filter((s) => s !== name) });
  };

  // ---------- Auth ----------
  const handleCreateFirstAdmin = async () => {
    setLoginError("");
    if (!setupName.trim() || !setupUsername.trim() || !setupPassword) {
      setLoginError("Please fill in all fields");
      return;
    }
    // The first account created becomes the main/admin account.
    // Only this account (or another account it later promotes) can manage employees.
    const admin = {
      name: setupName.trim(),
      username: setupUsername.trim(),
      password: await hashPassword(setupPassword),
      isAdmin: true,
    };
    await persistEmployees([admin]);
    await window.storage.set("tickets:setupComplete", "true", true).catch(() => {});
    setSetupComplete(true);
    await window.storage.set("session:user", admin.username, false);
    sessionStartedAtRef.current = Date.now();
    setCurrentUser({ username: admin.username, name: admin.name, isAdmin: true });
    setSetupName(""); setSetupUsername(""); setSetupPassword("");
  };

  const handleLogin = async () => {
    setLoginError("");
    if (loginLockUntil && Date.now() < loginLockUntil) {
      const secondsLeft = Math.ceil((loginLockUntil - Date.now()) / 1000);
      setLoginError(`Too many failed attempts. Try again in ${secondsLeft}s.`);
      return;
    }
    const candidate = (employees || []).find((e) => e.username === loginUsername.trim());
    const match = candidate && (await verifyPassword(candidate.password, loginPassword)) ? candidate : null;
    if (!match) {
      const nextCount = loginFailCount + 1;
      setLoginFailCount(nextCount);
      if (nextCount >= 5) {
        setLoginLockUntil(Date.now() + 30000);
        setLoginFailCount(0);
        setLoginError("Too many failed attempts. Try again in 30s.");
      } else {
        setLoginError("Incorrect username or password");
      }
      return;
    }
    setLoginFailCount(0);
    setLoginLockUntil(0);
    await window.storage.set("session:user", match.username, false);
    sessionStartedAtRef.current = Date.now();
    setCurrentUser({ username: match.username, name: match.name, isAdmin: !!match.isAdmin });
    setLoginUsername(""); setLoginPassword("");
    try {
      const lastSectionRes = await window.storage.get(`tickets:lastSection:${match.username}`, false).catch(() => null);
      const lastSection = lastSectionRes && lastSectionRes.value;
      if (["flights", "hotels", "visa", "cars", "files"].includes(lastSection)) {
        setActiveSection(lastSection);
      }
    } catch (e) {
      // Best-effort; falls back to the default "flights" section
    }
  };

  const handleLogout = async () => {
    await window.storage.delete("session:user", false).catch(() => {});
    setCurrentUser(null);
    setShowManage(false);
    setEditingUsername(null);
  };

  // Lets the main account remotely sign out any currently-online employee (or itself)
  // from the "online now" panel. This account has no way to reach into another browser's
  // own local session storage, so instead it writes a shared timestamped flag; that
  // employee's own client picks it up on its next heartbeat (every few seconds) and signs
  // itself out. Their presence is cleared immediately here so they show as offline right away.
  const handleForceSignOut = async (username) => {
    try {
      await window.storage.set(`tickets:forceLogout:${username}`, String(Date.now()), true);
      await window.storage.delete(`tickets:presence:${username}`, true).catch(() => {});
      setPresenceMap((prev) => {
        const next = { ...prev };
        delete next[username];
        return next;
      });
    } catch (e) {
      // Best-effort; the admin can just try again
    }
  };

  const handleAddEmployee = async () => {
    setManageError("");
    if (!currentUser.isAdmin && !isOwnerUser) {
      setManageError("Only the main account can add employees");
      return;
    }
    if (!newEmployee.name.trim() || !newEmployee.username.trim() || !newEmployee.password) {
      setManageError("Please fill in all fields");
      return;
    }
    if ((employees || []).some((e) => e.username === newEmployee.username.trim())) {
      setManageError("That username already exists");
      return;
    }
    const next = [
      ...(employees || []),
      {
        ...newEmployee,
        username: newEmployee.username.trim(),
        password: await hashPassword(newEmployee.password),
        isAdmin: false,
        ...reconcilePermissions(newEmployee),
      },
    ];
    await persistEmployees(next);
    setNewEmployee(emptyNewEmployee);
    setShowNewEmployeePerms(false);
  };

  // Applies a grade's preset permissions to an employee. The grade itself is stored
  // (for the badge/label), and every toggle it sets can still be flipped individually
  // afterwards via handleTogglePermission — the preset is just a fast starting point.
  const handleRoleChange = async (username, role) => {
    if (!currentUser.isAdmin && !isOwnerUser) {
      setManageError("Only the main account can change employee permissions");
      return;
    }
    const preset = ROLE_PRESETS[role] || ROLE_PRESETS.employee;
    const next = (employees || []).map((e) =>
      e.username === username ? { ...e, role, ...preset } : e
    );
    await persistEmployees(next);
  };

  // Single generic handler for every individual permission toggle (view all tickets,
  // add tickets, edit tickets, delete tickets, accounting/notes-only mode, manage
  // companies). Each toggle is independently switchable by the main account; coherence
  // between them (edit/delete requiring view, accounting overriding add/edit/delete) is
  // enforced afterwards by reconcilePermissions so the stored record never contradicts itself.
  const handleTogglePermission = async (username, field, checked) => {
    if (!currentUser.isAdmin && !isOwnerUser) {
      setManageError("Only the main account can change employee permissions");
      return;
    }
    const next = (employees || []).map((e) =>
      e.username === username ? { ...e, ...reconcilePermissions({ ...e, [field]: checked }) } : e
    );
    await persistEmployees(next);
  };

  // Toggles one section (Flights/Hotels/Visa/Transportation/Files) on or off for an
  // employee, independent of their ticket permissions above.
  const handleToggleSection = async (username, section, checked) => {
    if (!currentUser.isAdmin && !isOwnerUser) {
      setManageError("Only the main account can change employee permissions");
      return;
    }
    const next = (employees || []).map((e) =>
      e.username === username ? { ...e, sections: { ...employeeSections(e), [section]: checked } } : e
    );
    await persistEmployees(next);
  };

  // Sets one of View all services / Edit / Delete for one specific section (Flights,
  // Hotels, Visa, Transportation, or Files), independent of every other section's
  // permissions — and independent of the other two toggles within the same section too.
  const handleToggleSectionPermission = async (username, section, field, checked) => {
    if (!currentUser.isAdmin && !isOwnerUser) {
      setManageError("Only the main account can change employee permissions");
      return;
    }
    const next = (employees || []).map((e) => {
      if (e.username !== username) return e;
      const current = employeeSectionPerm(e, section);
      const updated = { ...current, [field]: checked };
      return { ...e, sectionPerms: { ...(e.sectionPerms || {}), [section]: updated } };
    });
    await persistEmployees(next);
  };

  // Promotes an employee to a main/admin account. Any main account can promote another one.
  const handlePromoteToAdmin = async (username) => {
    if (!currentUser.isAdmin) {
      setManageError("Only the main account can grant main-account access");
      return;
    }
    const target = (employees || []).find((e) => e.username === username);
    if (!target) return;
    requestConfirm(
      `Make "${target.name}" a main account? They will be able to manage all employees, permissions, backups, and see every ticket.`,
      async () => {
        const next = (employees || []).map((e) =>
          e.username === username ? { ...e, isAdmin: true } : e
        );
        await persistEmployees(next);
        setConfirmDialog(null);
      }
    );
  };

  // Demotes a main account back to a regular employee. Blocked if it would leave zero main accounts.
  const handleDemoteAdmin = async (username) => {
    if (!currentUser.isAdmin) {
      setManageError("Only the main account can remove main-account access");
      return;
    }
    const admins = (employees || []).filter((e) => e.isAdmin);
    if (admins.length <= 1) {
      setManageError("There must always be at least one main account");
      return;
    }
    const target = (employees || []).find((e) => e.username === username);
    if (!target) return;
    requestConfirm(`Remove main-account access from "${target.name}"?`, async () => {
      const next = (employees || []).map((e) =>
        e.username === username ? { ...e, isAdmin: false } : e
      );
      await persistEmployees(next);
      // If the admin demoted themselves, drop their manage-panel view since they're no longer main
      if (username === currentUser.username) {
        setCurrentUser({ ...currentUser, isAdmin: false });
        setShowManage(false);
      }
      setConfirmDialog(null);
    });
  };

  const handleDeleteEmployee = async (username) => {
    if (!currentUser.isAdmin && !isOwnerUser) {
      setManageError("Only the main account can remove employees");
      return;
    }
    if (username === currentUser.username) {
      setManageError("You can't delete the account you're logged in with");
      return;
    }
    await persistEmployees((employees || []).filter((e) => e.username !== username));
  };

  const startEditEmployee = (emp) => {
    setManageError("");
    setEditShowPassword(false);
    setEditingUsername(emp.username);
    setEditDraft({ name: emp.name, username: emp.username, password: emp.password });
  };

  const cancelEditEmployee = () => {
    setEditingUsername(null);
    setEditDraft({ name: "", username: "", password: "" });
  };

  const saveEditEmployee = async () => {
    if (!currentUser.isAdmin && !isOwnerUser) {
      setManageError("Only the main account can edit employee accounts");
      return;
    }
    // An Owner has admin-level access to everyone else, but must never be able to edit
    // a true main account's own credentials — that stays admin-to-admin only.
    const targetBeingEdited = (employees || []).find((e) => e.username === editingUsername);
    if (isOwnerUser && targetBeingEdited && targetBeingEdited.isAdmin) {
      setManageError("Only a main account can edit another main account");
      return;
    }
    setManageError("");
    const trimmedName = editDraft.name.trim();
    const trimmedUsername = editDraft.username.trim();
    if (!trimmedName || !trimmedUsername || !editDraft.password) {
      setManageError("Please fill in all fields");
      return;
    }
    const clash = (employees || []).some(
      (e) => e.username !== editingUsername && e.username === trimmedUsername
    );
    if (clash) {
      setManageError("That username is already taken by another account");
      return;
    }
    const targetPassword =
      editDraft.password === targetBeingEdited.password
        ? targetBeingEdited.password
        : await hashPassword(editDraft.password);
    const next = (employees || []).map((e) =>
      e.username === editingUsername
        ? { ...e, name: trimmedName, username: trimmedUsername, password: targetPassword }
        : e
    );
    await persistEmployees(next);

    // If the main account edited its own account, keep the current session in sync
    if (editingUsername === currentUser.username) {
      await window.storage.set("session:user", trimmedUsername, false);
      setCurrentUser({ ...currentUser, name: trimmedName, username: trimmedUsername });
    }
    cancelEditEmployee();
  };

  // Saves an employee's name/username/password from inside the Permissions modal.
  // Returns an error message on failure, or null on success — the modal shows the
  // error itself rather than relying on the page-level banner, which can sit behind it.
  const handleSaveEmployeeDetails = async (username, draft) => {
    if (!currentUser.isAdmin && !isOwnerUser) {
      return "Only the main account can edit employee accounts";
    }
    const target = (employees || []).find((e) => e.username === username);
    if (isOwnerUser && target && target.isAdmin) {
      return "Only a main account can edit another main account";
    }
    const trimmedName = draft.name.trim();
    const trimmedUsername = draft.username.trim();
    if (!trimmedName || !trimmedUsername || !draft.password) {
      return "Please fill in all fields";
    }
    const clash = (employees || []).some(
      (e) => e.username !== username && e.username === trimmedUsername
    );
    if (clash) {
      return "That username is already taken by another account";
    }
    const targetPassword =
      draft.password === target.password ? target.password : await hashPassword(draft.password);
    const next = (employees || []).map((e) =>
      e.username === username
        ? { ...e, name: trimmedName, username: trimmedUsername, password: targetPassword }
        : e
    );
    await persistEmployees(next);

    // If the main account edited its own account, keep the current session in sync
    if (username === currentUser.username) {
      await window.storage.set("session:user", trimmedUsername, false);
      setCurrentUser({ ...currentUser, name: trimmedName, username: trimmedUsername });
    }
    return null;
  };

  const handleChangePassword = async () => {
    setPasswordError("");
    setPasswordSuccess("");
    if (!currentPasswordInput || !newPasswordInput || !confirmPasswordInput) {
      setPasswordError("Please fill in all fields");
      return;
    }
    const me = (employees || []).find((e) => e.username === currentUser.username);
    if (!me || !(await verifyPassword(me.password, currentPasswordInput))) {
      setPasswordError("Current password is incorrect");
      return;
    }
    if (newPasswordInput !== confirmPasswordInput) {
      setPasswordError("New password and confirmation do not match");
      return;
    }
    if (newPasswordInput.length < 4) {
      setPasswordError("New password should be at least 4 characters");
      return;
    }
    const hashedNew = await hashPassword(newPasswordInput);
    const next = (employees || []).map((e) =>
      e.username === currentUser.username ? { ...e, password: hashedNew } : e
    );
    await persistEmployees(next);
    setPasswordSuccess("Password updated successfully");
    setCurrentPasswordInput("");
    setNewPasswordInput("");
    setConfirmPasswordInput("");
  };

  // ---------- Backup / restore (main account or Owner) ----------
  const handleBackup = () => {
    if (!currentUser.isAdmin && !isOwnerUser) return;
    const payload = {
      backupFormat: "flight-tickets-v1",
      exportedAt: new Date().toISOString(),
      exportedBy: currentUser.name,
      tickets,
      employees,
      suggestions,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `flight_tickets_backup_${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const triggerRestore = () => {
    if (!currentUser.isAdmin && !isOwnerUser) return;
    setRestoreError("");
    setRestoreSuccess("");
    fileInputRef.current && fileInputRef.current.click();
  };

  const handleRestoreFile = async (e) => {
    setRestoreError("");
    setRestoreSuccess("");
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      if (!parsed || !Array.isArray(parsed.tickets) || !Array.isArray(parsed.employees)) {
        setRestoreError("This file doesn't look like a valid backup");
        return;
      }
      // Normalize suggestions defensively so nothing from the backup is silently dropped,
      // even if the file is from an older/partial export.
      const s = parsed.suggestions || {};
      const normalizedSuggestions = {
        companies: Array.isArray(s.companies) ? s.companies : [],
        // Never restore saved customer names — this field must have no autocomplete history.
        customers: [],
        airlines: Array.isArray(s.airlines) ? s.airlines : [],
        cities: Array.isArray(s.cities) ? s.cities : [],
      };
      const suggestionsCount =
        normalizedSuggestions.companies.length +
        normalizedSuggestions.customers.length +
        normalizedSuggestions.airlines.length +
        normalizedSuggestions.cities.length;
      requestConfirm(
        "This will replace all current tickets and employee accounts with the data in this backup file. This cannot be undone. Continue?",
        async () => {
          await persistTickets(parsed.tickets);
          await persistEmployees(parsed.employees);
          await persistSuggestions(normalizedSuggestions);
          setRestoreSuccess(
            `Backup restored successfully: ${parsed.tickets.length} tickets, ${parsed.employees.length} employee accounts, and ${suggestionsCount} saved suggestions.`
          );
          setConfirmDialog(null);
        }
      );
    } catch (err) {
      setRestoreError("Could not read this backup file");
    } finally {
      e.target.value = "";
    }
  };

  // ---------- Tickets ----------
  // Builds a plain-language list of what changed between the ticket's previous version and
  // the edited one (e.g. "From: CAI → JED"), used to log every ticket edit — not just notes —
  // into the same edit-history trail, along with who made the change.
  const describeTicketChanges = (before, after) => {
    const changes = [];
    const fieldLabels = {
      company: "Company",
      supplier: "Supplier",
      from: "From",
      to: "To",
      airline: "Airline",
      date: "Date",
      netPrice: "Net price",
      soldPrice: "Sold price",
    };
    Object.keys(fieldLabels).forEach((key) => {
      const beforeVal = before[key] ?? "";
      const afterVal = after[key] ?? "";
      if (String(beforeVal) !== String(afterVal)) {
        changes.push(`${fieldLabels[key]}: ${beforeVal || "—"} → ${afterVal || "—"}`);
      }
    });

    const beforeCustomers = Array.isArray(before.customers) ? before.customers : [];
    const afterCustomers = Array.isArray(after.customers) ? after.customers : [];
    if (beforeCustomers.length !== afterCustomers.length) {
      changes.push(`Customers: ${beforeCustomers.length} → ${afterCustomers.length}`);
    }
    const maxLen = Math.max(beforeCustomers.length, afterCustomers.length);
    for (let i = 0; i < maxLen; i++) {
      const b = beforeCustomers[i] || { name: "", ticketNumber: "" };
      const a = afterCustomers[i] || { name: "", ticketNumber: "" };
      if ((b.name || "") !== (a.name || "")) {
        changes.push(`Customer ${i + 1} name: ${b.name || "—"} → ${a.name || "—"}`);
      }
      if ((b.ticketNumber || "") !== (a.ticketNumber || "")) {
        changes.push(`Customer ${i + 1} ticket number: ${b.ticketNumber || "—"} → ${a.ticketNumber || "—"}`);
      }
    }

    const beforeRefunds = getRefunds(before);
    const afterRefunds = getRefunds(after);
    if (beforeRefunds.length === 0 && afterRefunds.length > 0) {
      changes.push(`Refund added (${afterRefunds.length} ticket${afterRefunds.length > 1 ? "s" : ""})`);
    } else if (beforeRefunds.length > 0 && afterRefunds.length === 0) {
      changes.push("Refund removed");
    } else if (JSON.stringify(beforeRefunds) !== JSON.stringify(afterRefunds)) {
      changes.push("Refund updated");
    }
    return changes;
  };

  const handleSubmit = () => {
    setError("");
    const customers = form.customers || [];
    // A customer row normally needs a ticket number, but a filled-in PNR reference
    // covers the same purpose (identifying the booking), so either one satisfies
    // this check — the ticket number stops being mandatory once a PNR is entered.
    const customersValid =
      customers.length > 0 &&
      customers.every((c) => c.name.trim() && (c.ticketNumber.trim() || (c.pnrReference || "").trim()));
    // A multi-destination route needs at least two filled-in stops; a regular route
    // needs both From and To.
    const cleanDestinations = (form.destinations || []).map((d) => (d || "").trim()).filter(Boolean);
    const routeValid = form.multiDestination
      ? cleanDestinations.length >= 2
      : form.from.trim() && form.to.trim();
    if (!customersValid || !routeValid || form.netPrice === "" || form.soldPrice === "") {
      setError("Please enter at least the customer name(s), a ticket number or PNR reference for each, destinations, and prices");
      return;
    }
    // Keep the original owner when editing an existing ticket (so an admin editing someone
    // else's ticket doesn't reassign it to themselves); new tickets belong to whoever adds them.
    const isEditingExisting = !!(form.id && form.employeeUsername);
    const original = form.id ? tickets.find((t) => t.id === form.id) : null;
    let record = {
      ...form,
      customers,
      customersCount: customers.length,
      // For a multi-destination route, from/to are kept in sync as the first/last stop so
      // every place that reads a plain origin/destination (search, exports, older code)
      // keeps working; a regular route just keeps its own from/to untouched.
      destinations: form.multiDestination ? cleanDestinations : [],
      from: form.multiDestination ? cleanDestinations[0] || "" : form.from,
      to: form.multiDestination ? cleanDestinations[cleanDestinations.length - 1] || "" : form.to,
      // Return airport always mirrors the first (From) airport on a round trip — it's
      // not independently editable, so it's derived here rather than trusted from form state.
      returnAirport:
        form.tripType === "roundTrip"
          ? (form.multiDestination ? cleanDestinations[0] || "" : form.from)
          : "",
      employee: isEditingExisting ? form.employee : currentUser.name,
      employeeUsername: isEditingExisting ? form.employeeUsername : currentUser.username,
      id: form.id || Date.now().toString(),
    };
    // Every edit to an existing ticket — any field, not just notes — gets logged into the
    // same edit-history trail shown under Notes, recording what changed and who changed it.
    if (original) {
      const changes = describeTicketChanges(original, record);
      if (changes.length > 0) {
        const history = Array.isArray(original.notesHistory) ? original.notesHistory : [];
        record = {
          ...record,
          notesHistory: [
            ...history,
            { type: "edit", changes, by: currentUser.name, at: new Date().toISOString() },
          ],
        };
      }
    }
    let next;
    if (form.id) {
      next = tickets.map((t) => (t.id === form.id ? record : t));
    } else {
      next = [record, ...tickets];
    }
    persistTickets(next);
    rememberSuggestionsFromRecord(record);
    setForm(getEmptyForm());
    setSupplierOther(false);
  };

  // The main account can always edit tickets; an employee can too, but only if they've
  // been granted the "edit tickets" permission. Deleting stays main-account only either way.
  const handleEdit = (t) => {
    if (!currentUser.isAdmin && !canEditTickets) return;
    // Backward compatibility: older records stored a single customer/ticketNumber pair
    const customers =
      Array.isArray(t.customers) && t.customers.length > 0
        ? t.customers
        : [{ name: t.customer || "", ticketNumber: t.ticketNumber || "" }];
    // Backward compatibility: older records have no multiDestination/destinations fields.
    const destinations =
      Array.isArray(t.destinations) && t.destinations.length >= 2 ? t.destinations : [t.from || "", t.to || ""];
    setForm({ ...t, customers, customersCount: customers.length, multiDestination: !!t.multiDestination, destinations, tripType: t.tripType || "oneWay", returnAirport: t.returnAirport || "" });
    setSupplierOther(!!t.supplier && !SUPPLIERS.includes(t.supplier));
  };
  const handleDelete = (id) => {
    if (!currentUser.isAdmin && !canDeleteTickets) {
      setError("You don't have permission to delete tickets");
      return;
    }
    if (form.id === id) { setForm(getEmptyForm()); setSupplierOther(false); }
    persistTickets(tickets.filter((t) => t.id !== id));
  };
  const handleCancel = () => { setForm(getEmptyForm()); setSupplierOther(false); };

  // Opens the full-detail view ("page") for a ticket, showing every field including notes.
  const openTicketDetail = (t) => {
    setViewingTicketId(t.id);
    setNotesDraft(t.notes || "");
    setNotesSaved(false);
  };
  const closeTicketDetail = () => {
    setViewingTicketId(null);
    setNotesDraft("");
    setNotesSaved(false);
  };
  // Saves an edit to just the notes field of a ticket, without touching anything else.
  // Every save appends an entry to notesHistory recording who made the change and when,
  // so the full edit trail (including accounting-account edits) stays visible.
  const saveTicketNotes = (id) => {
    const now = new Date().toISOString();
    const nextNotes = notesDraft.toUpperCase();
    const next = tickets.map((t) => {
      if (t.id !== id) return t;
      const history = Array.isArray(t.notesHistory) ? t.notesHistory : [];
      return {
        ...t,
        notes: nextNotes,
        notesHistory: [...history, { value: nextNotes, by: currentUser.name, at: now }],
      };
    });
    persistTickets(next);
    setNotesSaved(true);
  };

  // Normalizes a ticket's refund records into a list. Current tickets store `refunds` as
  // an array — one entry per refunded customerIndex, since a single booking can have
  // several customers/tickets refunded independently. Older saved tickets may still have
  // a single `refund` object from before that change; treated here as a one-item list so
  // every reader below keeps working for both shapes without a separate migration step.
  const getRefunds = (t) => {
    if (!t) return [];
    if (Array.isArray(t.refunds)) return t.refunds;
    if (t.refund) return [t.refund];
    return [];
  };

  // True once at least one refund (either side) has actually been recorded for a ticket.
  const hasRefund = (t) => getRefunds(t).some((r) => r && (r.airlineAmount !== "" || r.customerAmount !== ""));

  // The recorded refund entry (if any) for one specific customer/ticket within a booking —
  // used to show the "Refunded" badge against the right customer row rather than every row.
  const refundForIndex = (t, i) =>
    getRefunds(t).find((r) => r && (r.customerIndex || 0) === i && (r.airlineAmount !== "" || r.customerAmount !== ""));

  // Accounting-adjusted figures for a ticket: every recorded refund is deducted from both
  // sides — what the airline paid back reduces our cost (net price), and what we paid
  // back to the customer reduces our revenue (sold price) — so sales/profit totals
  // everywhere (ticket rows, summary cards, monthly/company breakdowns, exports)
  // reflect the refund rather than the original pre-refund booking amounts.
  const netAfterRefund = (t) =>
    (parseFloat(t.netPrice) || 0) - getRefunds(t).reduce((sum, r) => sum + (parseFloat(r.airlineAmount) || 0), 0);
  const soldAfterRefund = (t) =>
    (parseFloat(t.soldPrice) || 0) - getRefunds(t).reduce((sum, r) => sum + (parseFloat(r.customerAmount) || 0), 0);
  const profitAfterRefund = (t) => soldAfterRefund(t) - netAfterRefund(t);

  const handleCustomersCountChange = (value) => {
    const count = value === "" ? "" : value;
    const customers = resizeCustomers(form.customers, value);
    // When more customer rows are added, auto-sequence their ticket numbers by
    // increasing the previous customer's number by one (only if it was filled in).
    // The PNR reference isn't sequenced like the ticket number — every passenger on
    // the same booking shares the same PNR — so new rows just inherit the first
    // customer's PNR reference verbatim (only if it was filled in).
    const firstPnr = form.customers[0] && form.customers[0].pnrReference;
    for (let i = form.customers.length; i < customers.length; i++) {
      const generated = nextTicketNumber(lastIssuedTicketNumber(customers[i - 1]));
      if (generated) customers[i] = { ...customers[i], ticketNumber: generated };
      if (firstPnr) customers[i] = { ...customers[i], pnrReference: firstPnr };
    }
    setForm({ ...form, customersCount: count, customers });
  };

  // From/To suggestions are shown as "CODE - City, Country" for easy searching, but only
  // the 3-letter IATA code should end up stored in the field/cell. If the typed or picked
  // value matches that "CODE - ..." shape, keep just the code; otherwise keep it as typed.
  const handleCityChange = (field, value) => {
    const raw = (value || "").toUpperCase();
    const match = raw.match(/^([A-Z]{3})\s*-\s*.+$/);
    setForm({ ...form, [field]: match ? match[1] : raw });
  };

  // Same "CODE - City, Country" → CODE cleanup as handleCityChange, but for one stop
  // in a multi-destination (multi-city) route.
  const handleDestinationChange = (index, value) => {
    const raw = (value || "").toUpperCase();
    const match = raw.match(/^([A-Z]{3})\s*-\s*.+$/);
    const clean = match ? match[1] : raw;
    const destinations = form.destinations.map((d, i) => (i === index ? clean : d));
    setForm({ ...form, destinations });
  };

  const addDestinationStop = () => {
    setForm({ ...form, destinations: [...form.destinations, ""] });
  };

  // Always keeps at least two stops (a route needs a start and an end).
  const removeDestinationStop = (index) => {
    const destinations = form.destinations.filter((_, i) => i !== index);
    setForm({ ...form, destinations: destinations.length >= 2 ? destinations : ["", ""] });
  };

  const handleAirlineChange = (value) => {
    const airline = value.toUpperCase();
    const code = getAirlineCodeByIata(airline);
    // If we recognize the airline code, pre-fill its 3-digit prefix into any customer's
    // ticket number that hasn't been typed into yet (never overwrites manual entries).
    const customers = code
      ? form.customers.map((c) => (c.ticketNumber ? c : { ...c, ticketNumber: `${code}-` }))
      : form.customers;
    setForm({ ...form, airline, customers });
  };

  const handleCustomerFieldChange = (index, field, value) => {
    let nextValue = (value || "").toUpperCase();
    if (field === "ticketNumber") {
      // Keep only letters and digits, then auto-insert a hyphen after the first 3 characters
      const clean = nextValue.replace(/[^A-Z0-9]/g, "").slice(0, 13);
      nextValue = clean.length > 3 ? `${clean.slice(0, 3)}-${clean.slice(3)}` : clean;
    } else if (field === "ticketNumber2") {
      // A conjunction ticket number is just the incremented 3-digit tail after a dash
      // (see conjunctionTicketSuffix) — not a full independent ticket number, so this
      // keeps only digits and re-applies the leading dash rather than the usual
      // prefix-then-hyphen formatting.
      const digits = nextValue.replace(/[^0-9]/g, "").slice(0, 3);
      nextValue = digits ? `-${digits}` : "";
    } else if (field === "pnrReference") {
      // PNR references are up to 6 letters/digits (the airline's booking locator).
      nextValue = nextValue.replace(/[^A-Z0-9]/g, "").slice(0, 6);
    }
    const customers = form.customers.map((c, i) => (i === index ? { ...c, [field]: nextValue } : c));
    let airline = form.airline;
    if (field === "ticketNumber") {
      // Auto-detect the airline from the ticket number's 3-digit prefix (only if the
      // airline field hasn't been filled in yet, so it never overrides a manual choice)
      if (!airline) {
        const match = nextValue.match(/^([A-Z0-9]{3})-/);
        if (match) {
          const detected = getAirlineByCode(match[1]);
          if (detected) airline = detected;
        }
      }
    }
    setForm({ ...form, customers, airline });
  };

  // Toggles whether a customer has a conjunction ticket (a second ticket number issued
  // together with their first). Checking it auto-fills the conjunction suffix from the
  // customer's first ticket number (still editable by hand afterward); unchecking clears
  // the second ticket number out.
  const handleCustomerConjunctionToggle = (index, checked) => {
    const customers = form.customers.map((c, i) =>
      i === index ? { ...c, conjunction: checked, ticketNumber2: checked ? conjunctionTicketSuffix(c.ticketNumber) : "" } : { ...c }
    );
    // Switching the conjunction on/off shifts where the sequence for later customers
    // should continue from, so re-run the same cascade as handleTicketNumberBlur below —
    // still stopping at the first customer whose ticket number is already filled in.
    let last = lastIssuedTicketNumber(customers[index]);
    for (let i = index + 1; i < customers.length; i++) {
      if (customers[i].ticketNumber) break;
      const generated = nextTicketNumber(last);
      if (!generated) break;
      customers[i] = { ...customers[i], ticketNumber: generated };
      last = generated;
    }
    setForm({ ...form, customers });
  };

  // Runs once the person leaves the ticket number field (not on every keystroke), using
  // whatever they finished typing, and auto-fills any following ticket numbers that are
  // still empty — each one increasing the previous by one. Stops at the first one someone
  // has already typed something into, so manual entries are never overwritten. Also keeps
  // this customer's conjunction suffix (if any) in sync with their first ticket number.
  const handleTicketNumberBlur = (index) => {
    const customers = form.customers.map((c) => ({ ...c }));
    let last = customers[index] && customers[index].ticketNumber;
    if (!last) return;
    if (customers[index].conjunction) {
      customers[index].ticketNumber2 = conjunctionTicketSuffix(last);
    }
    // If this customer has a conjunction ticket, that second number was already issued
    // to them — continue the sequence for later customers after ITS tail, not the first
    // ticket's tail.
    last = lastIssuedTicketNumber(customers[index]);
    for (let i = index + 1; i < customers.length; i++) {
      if (customers[i].ticketNumber) break;
      const generated = nextTicketNumber(last);
      if (!generated) break;
      customers[i] = { ...customers[i], ticketNumber: generated };
      last = generated;
    }
    setForm({ ...form, customers });
  };

  // Runs once the person leaves the PNR reference field. Only the FIRST customer's PNR
  // drives the rest — all passengers on the same booking normally share one PNR — so
  // finishing typing it there copies it into every other customer row that's still
  // empty. Rows already typed into by hand (e.g. a different PNR for an interline
  // passenger) are left alone, and each ticket number stays independently editable —
  // the PNR reference is a separate field and never touches it.
  const handlePnrReferenceBlur = (index) => {
    if (index !== 0) return;
    const value = form.customers[0] && form.customers[0].pnrReference;
    if (!value) return;
    const customers = form.customers.map((c, i) => (i === 0 || c.pnrReference ? c : { ...c, pnrReference: value }));
    setForm({ ...form, customers });
  };

  // Finds a saved ticket by ticket number, searching every customer row across all
  // saved tickets (old or current schema). Used when a reissued ticket references an
  // older one, both to auto-fill its issue date and to import the rest of its data.
  const findTicketByNumber = (ticketNumber) => {
    const target = (ticketNumber || "").trim().toUpperCase();
    if (!target) return null;
    for (const t of tickets) {
      const custs =
        Array.isArray(t.customers) && t.customers.length > 0
          ? t.customers
          : [{ name: t.customer || "", ticketNumber: t.ticketNumber || "" }];
      if (
        custs.some(
          (c) =>
            (c.ticketNumber || "").trim().toUpperCase() === target ||
            (c.ticketNumber2 || "").trim().toUpperCase() === target
        )
      ) {
        return t;
      }
    }
    return null;
  };

  // Cleans up the old ticket number the same way regular ticket numbers are formatted.
  const handleOldTicketNumberChange = (value) => {
    const clean = (value || "").toUpperCase().replace(/[^A-Z0-9]/g, "").slice(0, 13);
    const nextValue = clean.length > 3 ? `${clean.slice(0, 3)}-${clean.slice(3)}` : clean;
    setForm({ ...form, oldTicketNumber: nextValue, oldTicketIssueDate: "" });
  };

  // Once the person finishes typing the old ticket number, look it up against saved
  // tickets and import that old ticket's data into the reissue form: issue date, company,
  // supplier, route (including a multi-destination route), airline, prices, and any
  // customer name not already typed. Anything the person already entered by hand is left
  // untouched — this only fills in fields that are still empty.
  const handleOldTicketNumberBlur = () => {
    const oldTicket = findTicketByNumber(form.oldTicketNumber);
    if (!oldTicket) {
      setForm({ ...form, oldTicketIssueDate: "" });
      return;
    }
    const oldCustomers =
      Array.isArray(oldTicket.customers) && oldTicket.customers.length > 0
        ? oldTicket.customers
        : [{ name: oldTicket.customer || "", ticketNumber: oldTicket.ticketNumber || "" }];
    // Fill in any customer row that doesn't have a name yet with the matching old
    // customer's name (by position); new ticket numbers are always left exactly as typed.
    const customers = form.customers.map((c, i) =>
      c.name.trim() ? c : { ...c, name: (oldCustomers[i] && oldCustomers[i].name) || c.name }
    );
    const hasOwnDestinations = (form.destinations || []).some((d) => (d || "").trim());
    setForm({
      ...form,
      oldTicketIssueDate: oldTicket.date || "",
      company: form.company || oldTicket.company || "",
      from: form.from || oldTicket.from || "",
      to: form.to || oldTicket.to || "",
      multiDestination: form.multiDestination || !!oldTicket.multiDestination,
      destinations: hasOwnDestinations
        ? form.destinations
        : Array.isArray(oldTicket.destinations) && oldTicket.destinations.length >= 2
        ? oldTicket.destinations
        : form.destinations,
      airline: form.airline || oldTicket.airline || "",
      netPrice: form.netPrice !== "" ? form.netPrice : oldTicket.netPrice ?? "",
      soldPrice: form.soldPrice !== "" ? form.soldPrice : oldTicket.soldPrice ?? "",
      customers,
    });
  };

  // Cleans up a refund row's ticket number the same way regular ticket numbers and the
  // reissue lookup are formatted.
  const handleRefundRowNumberChange = (index, value) => {
    const clean = (value || "").toUpperCase().replace(/[^A-Z0-9]/g, "").slice(0, 13);
    const nextValue = clean.length > 3 ? `${clean.slice(0, 3)}-${clean.slice(3)}` : clean;
    setRefundRows(refundRows.map((r, i) => (i === index ? { ...r, number: nextValue } : r)));
    setRefundSaved(false);
  };

  // Once the person finishes typing a row's ticket number, look it up and, if it already
  // has a refund recorded for the default (first) customer, load those amounts in for
  // editing; otherwise start blank. Picking a different customer below reloads that
  // customer's own recorded refund, if any — see the "Refunded ticket" select below.
  const handleRefundRowNumberBlur = (index) => {
    const target = findTicketByNumber(refundRows[index].number);
    const existing = target ? getRefunds(target).find((r) => (r.customerIndex || 0) === 0) : null;
    setRefundRows(
      refundRows.map((r, i) =>
        i === index
          ? {
              ...r,
              airlineAmount: existing ? existing.airlineAmount || "" : "",
              customerAmount: existing ? existing.customerAmount || "" : "",
              customerIndex: existing ? existing.customerIndex || 0 : 0,
            }
          : r
      )
    );
  };

  const addRefundRow = () => {
    setRefundRows([...refundRows, { number: "", airlineAmount: "", customerAmount: "", customerIndex: 0 }]);
  };

  const removeRefundRow = (index) => {
    setRefundRows(refundRows.length > 1 ? refundRows.filter((_, i) => i !== index) : refundRows);
  };

  // Saves every row's refund directly onto whichever saved ticket matches its typed
  // ticket number — independent of whatever ticket the main form is currently
  // adding/editing. Rows with no matching ticket are skipped. Two or more rows can point
  // at the same booking (e.g. refunding several customers on one multi-passenger ticket
  // record) — those are grouped and merged in by customerIndex rather than one row
  // overwriting another, and any of that booking's other already-recorded refunds (for
  // customers not touched by this save) are kept untouched. Logged into each affected
  // ticket's own edit-history trail.
  const saveAllRefunds = () => {
    const now = new Date().toISOString();
    const rowsByTicketId = {};
    refundRows.forEach((row) => {
      const target = findTicketByNumber(row.number);
      if (!target) return;
      const customerIndex = row.customerIndex || 0;
      if (!rowsByTicketId[target.id]) rowsByTicketId[target.id] = {};
      rowsByTicketId[target.id][customerIndex] = {
        airlineAmount: row.airlineAmount,
        customerAmount: row.customerAmount,
        customerIndex,
        date: todayDateStr(),
      };
    });
    if (Object.keys(rowsByTicketId).length === 0) return;
    const next = tickets.map((t) => {
      const newByIndex = rowsByTicketId[t.id];
      if (!newByIndex) return t;
      const newEntries = Object.values(newByIndex);
      const untouched = getRefunds(t).filter((r) => !((r.customerIndex || 0) in newByIndex));
      const history = Array.isArray(t.notesHistory) ? t.notesHistory : [];
      const summary = newEntries
        .map((r) => `airline ${r.airlineAmount || 0}, customer ${r.customerAmount || 0}`)
        .join("; ");
      return {
        ...t,
        refund: null,
        refunds: [...untouched, ...newEntries],
        notesHistory: [
          ...history,
          { type: "edit", changes: [`Refund: ${summary}`], by: currentUser.name, at: now },
        ],
      };
    });
    persistTickets(next);
    setRefundSaved(true);
  };

  // Removes only the specific customer/ticket refunds represented by the currently typed
  // rows (e.g. when switching away from the refund option, or unchecking it) — leaving any
  // other refund already recorded on the same booking, for a different customer, in place.
  // Keeps an entry in each affected ticket's edit-history trail.
  const clearAllRefundRows = () => {
    const now = new Date().toISOString();
    const indexesByTicketId = {};
    refundRows.forEach((row) => {
      const target = findTicketByNumber(row.number);
      if (!target) return;
      if (!indexesByTicketId[target.id]) indexesByTicketId[target.id] = new Set();
      indexesByTicketId[target.id].add(row.customerIndex || 0);
    });
    if (Object.keys(indexesByTicketId).length === 0) return;
    const next = tickets.map((t) => {
      const indexesToClear = indexesByTicketId[t.id];
      if (!indexesToClear) return t;
      const existing = getRefunds(t);
      const remaining = existing.filter((r) => !indexesToClear.has(r.customerIndex || 0));
      if (remaining.length === existing.length) return t;
      const history = Array.isArray(t.notesHistory) ? t.notesHistory : [];
      return {
        ...t,
        refund: null,
        refunds: remaining,
        notesHistory: [...history, { type: "edit", changes: ["Refund removed"], by: currentUser.name, at: now }],
      };
    });
    persistTickets(next);
  };

  // The main account always sees everything; employees see only what they entered,
  // unless the main account has granted them permission to view all tickets — or granted
  // them permission to edit tickets, since editing every ticket requires seeing every ticket.
  // Guarded against currentUser being null (e.g. on the login/setup screens).
  const currentEmployeeRecord = currentUser
    ? (employees || []).find((e) => e.username === currentUser.username)
    : null;
  // The main account always has every section; everyone else is gated by their
  // individually-granted section access (defaulting to all-allowed for legacy records).
  const mySections = currentUser && currentUser.isAdmin ? DEFAULT_SECTIONS : employeeSections(currentEmployeeRecord);
  // Accounts is a separate, financially-sensitive section that isn't part of the
  // per-employee sections/ownership system above — it's reserved for the main account
  // and any employee on the Owner/GM/Accountant grades (isOwner or isAccounting).
  const canAccessAccounts =
    !!currentUser &&
    (currentUser.isAdmin || !!(currentEmployeeRecord && (currentEmployeeRecord.isOwner || currentEmployeeRecord.isAccounting)));
  // If the current section is no longer (or was never) allowed for this employee —
  // e.g. their access was just changed by the main account — bounce them to the first
  // section they do have access to, instead of leaving them stuck on a blocked one.
  // Accounts is skipped here since it's gated by canAccessAccounts, not mySections.
  useEffect(() => {
    if (!currentUser) return;
    if (activeSection === "accounts") return;
    if (mySections[activeSection]) return;
    const firstAllowed = SECTION_OPTIONS.find((s) => mySections[s.value]);
    if (firstAllowed) setActiveSection(firstAllowed.value);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentUser, activeSection, mySections.flights, mySections.hotels, mySections.visa, mySections.cars, mySections.files]);
  useEffect(() => {
    if (!currentUser) return;
    if (activeSection === "accounts" && !canAccessAccounts) {
      const firstAllowed = SECTION_OPTIONS.find((s) => mySections[s.value]);
      if (firstAllowed) setActiveSection(firstAllowed.value);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentUser, activeSection, canAccessAccounts]);
  // View all / Edit / Delete, resolved independently for one section at a time. The main
  // account always gets everything; an accounting account gets view-all everywhere but
  // never edit/delete (their only allowed edit anywhere is the Notes field); everyone
  // else gets whatever was individually granted for that specific section (see
  // employeeSectionPerm — falls back to their old account-wide toggles if a section was
  // never customized).
  const sectionPermFor = (section) => {
    if (!currentUser) return { canViewAll: false, canEdit: false, canDelete: false };
    if (currentUser.isAdmin) return { canViewAll: true, canEdit: true, canDelete: true };
    return employeeSectionPerm(currentEmployeeRecord, section);
  };
  const flightsPerm = sectionPermFor("flights");
  const hotelsPerm = sectionPermFor("hotels");
  const visaPerm = sectionPermFor("visa");
  const carsPerm = sectionPermFor("cars");
  const filesPerm = sectionPermFor("files");
  // The Flights section is the app's original/default section, so the handful of
  // generic (non-section-scoped) helpers below — the flight ticket handlers, the
  // top-header badge, and the flight ticket detail page — keep using these names.
  const canViewAllTickets = flightsPerm.canViewAll;
  const canEditTickets = flightsPerm.canEdit;
  const canDeleteTickets = flightsPerm.canDelete;
  // Accounting accounts can see everything but cannot add tickets — their only allowed
  // edit anywhere in the app is the Notes field on a ticket's detail page.
  const isAccountingUser =
    !!currentUser && !currentUser.isAdmin && !!(currentEmployeeRecord && currentEmployeeRecord.isAccounting);
  // Every employee can add new tickets — this is no longer an individually
  // switchable permission. Accounting accounts are the one exception: their only
  // allowed edit anywhere in the app is the Notes field.
  const canAddTickets =
    !!currentUser &&
    (currentUser.isAdmin ||
      !!(currentEmployeeRecord && !currentEmployeeRecord.isAccounting));
  // A separate permission axis from ticket access: whether this account can add/edit/
  // remove saved company records (name, tax number, commercial register, phone numbers).
  const canManageCompanies =
    !!currentUser &&
    (currentUser.isAdmin || !!(currentEmployeeRecord && currentEmployeeRecord.canManageCompanies));
  // A step above the other toggles: an Owner-grade employee gets admin-level access to
  // Manage employees and Backup/Restore, but never the License panel — that stays
  // reserved for true main accounts (currentUser.isAdmin) so an Owner can never grant
  // themselves (or anyone else) admin access and route around this restriction.
  const isOwnerUser =
    !!currentUser && !currentUser.isAdmin && !!(currentEmployeeRecord && currentEmployeeRecord.isOwner);
  const hasAdminAccess = !!currentUser && (currentUser.isAdmin || isOwnerUser);
  const myPendingRequestsCount = (requests || []).filter(
    (r) => currentUser && r.toUsername === currentUser.username && r.status === "pending"
  ).length;
  // An Owner should never see that a main/admin account exists at all, so admin usernames
  // are dropped from the online-presence list whenever the viewer isn't a true admin.
  const visibleOnlineUsernames = onlineUsernames.filter((u) => {
    if (!currentUser || currentUser.isAdmin) return true;
    const emp = (employees || []).find((e) => e.username === u);
    return !(emp && emp.isAdmin);
  });
  // Roster shown in the header banner: every employee account, with whoever is
  // currently online sorted to the top (then alphabetically within each group).
  const employeeRoster = (employees || [])
    .filter((e) => currentUser && (currentUser.isAdmin || !e.isAdmin))
    .slice()
    .sort((a, b) => {
      const aOnline = onlineUsernames.includes(a.username) ? 0 : 1;
      const bOnline = onlineUsernames.includes(b.username) ? 0 : 1;
      if (aOnline !== bOnline) return aOnline - bOnline;
      return (a.name || "").localeCompare(b.name || "");
    });
  const visibleTickets = !currentUser
    ? []
    : canViewAllTickets
    ? tickets
    : tickets.filter((t) =>
        t.employeeUsername ? t.employeeUsername === currentUser.username : t.employee === currentUser.name
      );

  const visibleHotelBookings = !currentUser
    ? []
    : hotelsPerm.canViewAll
    ? hotelBookings
    : hotelBookings.filter((h) =>
        h.employeeUsername ? h.employeeUsername === currentUser.username : h.employee === currentUser.name
      );

  // Number of nights a single date range covers, from check-in to check-out (at least 1).
  const nightsBetween = (checkIn, checkOut) => {
    if (!checkIn || !checkOut) return 1;
    const inD = new Date(checkIn);
    const outD = new Date(checkOut);
    const diffDays = Math.round((outD - inD) / (1000 * 60 * 60 * 24));
    return diffDays > 0 ? diffDays : 1;
  };
  // Nights for one room line. Falls back to the booking's own (legacy) check-in/check-out
  // if the line itself doesn't have dates — older bookings saved before dates lived on
  // each room line.
  const roomLineNights = (l, h) => nightsBetween(l.checkIn || (h && h.checkIn), l.checkOut || (h && h.checkOut));
  // The overall date range shown for a booking: earliest check-in to latest check-out
  // across all its room lines.
  const hotelDateRange = (h) => {
    const lines = h.roomLines || [];
    const checkIns = lines.map((l) => l.checkIn || h.checkIn).filter(Boolean);
    const checkOuts = lines.map((l) => l.checkOut || h.checkOut).filter(Boolean);
    if (checkIns.length === 0 || checkOuts.length === 0) return { start: "", end: "" };
    return {
      start: checkIns.reduce((a, b) => (a < b ? a : b)),
      end: checkOuts.reduce((a, b) => (a > b ? a : b)),
    };
  };

  // Converts an amount from a room line's own currency into EGP, using the entered
  // USD->EGP rate. Returns the amount unchanged for EGP-priced lines.
  const hotelInEgp = (amount, currency) => (currency === "USD" ? amount * (usdToEgpRate || 0) : amount);

  // Per-booking totals: each room line's net/sold price is multiplied by its own room
  // count and its own number of nights, then summed across every line (e.g. 1 single
  // + 2 doubles, each possibly with different dates and currencies, all converted into
  // EGP to total).
  const hotelRoomCount = (h) => (h.roomLines || []).reduce((sum, l) => sum + (parseInt(l.count, 10) || 0), 0);
  // Raw (un-converted, in the line's own currency) total for one line — used when showing
  // a line's own subtotal next to its own currency in the form.
  const hotelLineNetTotal = (l, nights) => (parseFloat(l.netPrice) || 0) * (parseInt(l.count, 10) || 0) * nights;
  const hotelLineSoldTotal = (l, nights) => (parseFloat(l.soldPrice) || 0) * (parseInt(l.count, 10) || 0) * nights;
  const hotelNetTotal = (h) =>
    (h.roomLines || []).reduce((sum, l) => sum + hotelInEgp(hotelLineNetTotal(l, roomLineNights(l, h)), l.currency), 0);
  const hotelSoldTotal = (h) =>
    (h.roomLines || []).reduce((sum, l) => sum + hotelInEgp(hotelLineSoldTotal(l, roomLineNights(l, h)), l.currency), 0);
  const hotelProfitTotal = (h) => hotelSoldTotal(h) - hotelNetTotal(h);

  // Visa prices are entered per applicant, so a booking's real net/sold amounts are the
  // per-person price multiplied by how many customers are on that booking (falls back to
  // 1 if the customer list is empty, so older records without a list still total correctly).
  const visaCustomersCount = (v) => (v.customers || []).length || 1;
  const visaNetTotal = (v) => (parseFloat(v.netPrice) || 0) * visaCustomersCount(v);
  const visaSoldTotal = (v) => (parseFloat(v.soldPrice) || 0) * visaCustomersCount(v);
  const visaProfitTotal = (v) => visaSoldTotal(v) - visaNetTotal(v);
  // A booking is Corporate when a company name was entered; otherwise it's an
  // Individual booking automatically — no separate toggle needed.
  const hotelBookingType = (h) => (h.customer && h.customer.trim() ? "Corporate" : "Individual");
  // A short readable summary of a booking's room lines, e.g. "1x Single (BB, EGP, 01-AUG-2026→05-AUG-2026), 2x Double (AI, USD, 01-AUG-2026→03-AUG-2026)".
  const hotelLinesSummary = (h) =>
    (h.roomLines || [])
      .map((l) => {
        const type = ROOM_TYPES.find((r) => r.value === l.roomType)?.label || l.roomType;
        const meal = MEAL_PLANS.find((m) => m.value === l.mealPlan)?.value.toUpperCase() || "";
        const checkIn = l.checkIn || h.checkIn;
        const checkOut = l.checkOut || h.checkOut;
        const dates = checkIn && checkOut ? `, ${formatDisplayDate(checkIn)}→${formatDisplayDate(checkOut)}` : "";
        return `${l.count}× ${type} (${meal}, ${l.currency}${dates})`;
      })
      .join(", ");

  // ---------- Hotels: search + filters ----------
  const hotelMonthsAvailable = Array.from(
    new Set(visibleHotelBookings.map((h) => monthKey(h.bookingDate)))
  ).sort((a, b) => b.localeCompare(a));
  const hotelYearsAvailable = Array.from(
    new Set(visibleHotelBookings.map((h) => (h.bookingDate ? h.bookingDate.slice(0, 4) : "")).filter(Boolean))
  ).sort((a, b) => b.localeCompare(a));
  const hotelEmployeesAvailable = Array.from(
    new Set(visibleHotelBookings.map((h) => (h.employee || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));
  const hotelSuppliersAvailable = Array.from(
    new Set(visibleHotelBookings.map((h) => (h.supplier || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));
  const hotelNamesAvailable = Array.from(
    new Set(visibleHotelBookings.map((h) => (h.hotel || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));

  const hasActiveHotelFilter = !!(
    hotelSelectedYear.length || hotelSelectedMonth.length || hotelSelectedEmployee.length || hotelSelectedSupplier.length || hotelSelectedHotelName.length || hotelQuery.trim()
  );
  const activeHotelFilterCount =
    hotelSelectedYear.length + hotelSelectedMonth.length + hotelSelectedEmployee.length + hotelSelectedSupplier.length + hotelSelectedHotelName.length + (hotelQuery.trim() ? 1 : 0);
  const clearAllHotelFilters = () => {
    setHotelQuery("");
    setHotelSelectedYear([]);
    setHotelSelectedMonth([]);
    setHotelSelectedEmployee([]);
    setHotelSelectedSupplier([]);
    setHotelSelectedHotelName([]);
  };
  const filteredHotelBookings = visibleHotelBookings.filter((h) => {
    if (hotelSelectedYear.length && !hotelSelectedYear.includes((h.bookingDate || "").slice(0, 4))) return false;
    if (hotelSelectedMonth.length && !hotelSelectedMonth.includes(monthKey(h.bookingDate))) return false;
    if (hotelSelectedEmployee.length && !hotelSelectedEmployee.includes((h.employee || "").trim())) return false;
    if (hotelSelectedSupplier.length && !hotelSelectedSupplier.includes((h.supplier || "").trim())) return false;
    if (hotelSelectedHotelName.length && !hotelSelectedHotelName.includes((h.hotel || "").trim())) return false;
    const q = hotelQuery.trim().toLowerCase();
    if (!q) return true;
    return (
      (h.employee || "").toLowerCase().includes(q) ||
      (h.customer || "").toLowerCase().includes(q) ||
      (h.hotel || "").toLowerCase().includes(q) ||
      (h.supplier || "").toLowerCase().includes(q) ||
      (h.notes || "").toLowerCase().includes(q)
    );
  });

  // ---------- Visa: search + filters ----------
  // Visa bookings are tagged with an owning employee (employeeUsername), same as
  // Flights/Hotels/Files — falls back to matching on the display name for any legacy
  // booking saved before this field existed. There's still no Employee filter dropdown
  // here, unlike Hotels, since that wasn't asked for.
  const visibleVisaBookings = !currentUser
    ? []
    : visaPerm.canViewAll
    ? visaBookings
    : visaBookings.filter((v) =>
        v.employeeUsername ? v.employeeUsername === currentUser.username : v.employee === currentUser.name
      );
  const visaMonthsAvailable = Array.from(
    new Set(visibleVisaBookings.map((v) => monthKey(v.bookingDate)))
  ).sort((a, b) => b.localeCompare(a));
  const visaYearsAvailable = Array.from(
    new Set(visibleVisaBookings.map((v) => (v.bookingDate ? v.bookingDate.slice(0, 4) : "")).filter(Boolean))
  ).sort((a, b) => b.localeCompare(a));
  const visaSuppliersAvailable = Array.from(
    new Set(visibleVisaBookings.map((v) => (v.supplier || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));

  const hasActiveVisaFilter = !!(visaSelectedYear.length || visaSelectedMonth.length || visaSelectedSupplier.length || visaQuery.trim());
  const activeVisaFilterCount = visaSelectedYear.length + visaSelectedMonth.length + visaSelectedSupplier.length + (visaQuery.trim() ? 1 : 0);
  const clearAllVisaFilters = () => {
    setVisaQuery("");
    setVisaSelectedYear([]);
    setVisaSelectedMonth([]);
    setVisaSelectedSupplier([]);
  };
  const filteredVisaBookings = visibleVisaBookings.filter((v) => {
    if (visaSelectedYear.length && !visaSelectedYear.includes((v.bookingDate || "").slice(0, 4))) return false;
    if (visaSelectedMonth.length && !visaSelectedMonth.includes(monthKey(v.bookingDate))) return false;
    if (visaSelectedSupplier.length && !visaSelectedSupplier.includes((v.supplier || "").trim())) return false;
    const q = visaQuery.trim().toLowerCase();
    if (!q) return true;
    const customerNames = (v.customers || []).map((c) => c.name || "").join(" ");
    return (
      (v.visaType || "").toLowerCase().includes(q) ||
      (v.supplier || "").toLowerCase().includes(q) ||
      customerNames.toLowerCase().includes(q)
    );
  });

  // ---------- Transportation (cars): search + filters ----------
  // Car bookings are tagged with an owning employee (employeeUsername) the same way,
  // falling back to the display name for any legacy booking saved before this field
  // existed. Still no Employee filter dropdown here, unlike Hotels.
  const visibleCarBookings = !currentUser
    ? []
    : carsPerm.canViewAll
    ? carBookings
    : carBookings.filter((c) =>
        c.employeeUsername ? c.employeeUsername === currentUser.username : c.employee === currentUser.name
      );
  const carMonthsAvailable = Array.from(
    new Set(visibleCarBookings.map((c) => monthKey(c.bookingDate)))
  ).sort((a, b) => b.localeCompare(a));
  const carYearsAvailable = Array.from(
    new Set(visibleCarBookings.map((c) => (c.bookingDate ? c.bookingDate.slice(0, 4) : "")).filter(Boolean))
  ).sort((a, b) => b.localeCompare(a));
  const carSuppliersAvailable = Array.from(
    new Set(visibleCarBookings.map((c) => (c.supplier || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));

  const hasActiveCarFilter = !!(carSelectedYear.length || carSelectedMonth.length || carSelectedSupplier.length || carQuery.trim());
  const activeCarFilterCount = carSelectedYear.length + carSelectedMonth.length + carSelectedSupplier.length + (carQuery.trim() ? 1 : 0);
  const clearAllCarFilters = () => {
    setCarQuery("");
    setCarSelectedYear([]);
    setCarSelectedMonth([]);
    setCarSelectedSupplier([]);
  };
  const filteredCarBookings = visibleCarBookings.filter((c) => {
    if (carSelectedYear.length && !carSelectedYear.includes((c.bookingDate || "").slice(0, 4))) return false;
    if (carSelectedMonth.length && !carSelectedMonth.includes(monthKey(c.bookingDate))) return false;
    if (carSelectedSupplier.length && !carSelectedSupplier.includes((c.supplier || "").trim())) return false;
    const q = carQuery.trim().toLowerCase();
    if (!q) return true;
    return (
      (c.customerName || "").toLowerCase().includes(q) ||
      (c.routeFrom || "").toLowerCase().includes(q) ||
      (c.routeTo || "").toLowerCase().includes(q) ||
      (c.carType || "").toLowerCase().includes(q) ||
      (c.supplier || "").toLowerCase().includes(q) ||
      (c.flightNumber || "").toLowerCase().includes(q)
    );
  });

  const hotelTotals = filteredHotelBookings.reduce(
    (acc, h) => {
      acc.count += 1;
      acc.net += hotelNetTotal(h);
      acc.sold += hotelSoldTotal(h);
      acc.profit += hotelProfitTotal(h);
      return acc;
    },
    { count: 0, net: 0, sold: 0, profit: 0 }
  );

  // Visa and Transfers totals, same EGP-conversion approach as hotelTotals above
  // (each booking's own currency is converted to EGP so mixed-currency bookings can
  // be summed together). Counts the number of applicants/customers on each visa
  // booking, and the number of bookings for transfers.
  const visaTotals = filteredVisaBookings.reduce(
    (acc, v) => {
      const net = hotelInEgp(visaNetTotal(v), v.currency);
      const sold = hotelInEgp(visaSoldTotal(v), v.currency);
      acc.count += visaCustomersCount(v);
      acc.net += net;
      acc.sold += sold;
      acc.profit += sold - net;
      return acc;
    },
    { count: 0, net: 0, sold: 0, profit: 0 }
  );
  const carTotals = filteredCarBookings.reduce(
    (acc, c) => {
      const net = hotelInEgp(parseFloat(c.netPrice) || 0, c.currency);
      const sold = hotelInEgp(parseFloat(c.soldPrice) || 0, c.currency);
      acc.count += 1;
      acc.net += net;
      acc.sold += sold;
      acc.profit += sold - net;
      return acc;
    },
    { count: 0, net: 0, sold: 0, profit: 0 }
  );


  const getCustomers = (t) =>
    Array.isArray(t.customers) && t.customers.length > 0
      ? t.customers
      : [{ name: t.customer || "", ticketNumber: t.ticketNumber || "" }];

  // Visa bookings are now tagged with an owning employee like every other section, so
  // the "copy to a file" picker in Files only offers bookings this employee can already
  // see in the main Visa list — reusing the same permission-filtered list.
  const visibleVisaBookingsForFiles = visibleVisaBookings;

  // ============================================================
  // ---------- Accounts (accounting module) ----------
  // ============================================================
  // A flattened, EGP-normalized view of every booking across the four operational
  // sections (ALL of them, not just the currently search-filtered/permission-visible
  // lists above — the accounting module always needs the full financial picture).
  // Each entry's net/sold is converted to EGP using the same hotelInEgp() convention
  // already used for the cross-section totals above, so a supplier or customer who
  // appears in more than one section (and currency) can be combined into one figure.
  const acctBookings = [
    ...tickets.map((t) => ({
      key: `flights-${t.id}`,
      section: "flights",
      supplier: (t.supplier || "").trim(),
      customers: getCustomers(t).map((c) => (c.name || "").trim()).filter(Boolean),
      date: t.date || "",
      net: netAfterRefund(t),
      sold: soldAfterRefund(t),
    })),
    ...hotelBookings.map((h) => ({
      key: `hotels-${h.id}`,
      section: "hotels",
      supplier: (h.supplier || "").trim(),
      customers: [(h.customer || "").trim()].filter(Boolean),
      date: h.bookingDate || "",
      net: hotelNetTotal(h),
      sold: hotelSoldTotal(h),
    })),
    ...visaBookings.map((v) => ({
      key: `visa-${v.id}`,
      section: "visa",
      supplier: (v.supplier || "").trim(),
      customers: (v.customers || []).map((c) => (c.name || "").trim()).filter(Boolean),
      date: v.bookingDate || "",
      net: hotelInEgp(visaNetTotal(v), v.currency),
      sold: hotelInEgp(visaSoldTotal(v), v.currency),
    })),
    ...carBookings.map((c) => ({
      key: `cars-${c.id}`,
      section: "cars",
      supplier: (c.supplier || "").trim(),
      customers: [(c.customerName || "").trim()].filter(Boolean),
      date: c.bookingDate || "",
      net: hotelInEgp(parseFloat(c.netPrice) || 0, c.currency),
      sold: hotelInEgp(parseFloat(c.soldPrice) || 0, c.currency),
    })),
  ];

  const SECTION_LABELS_AR = { flights: "طيران", hotels: "فنادق", visa: "فيزا", cars: "ترانسفير" };
  const SECTION_LABELS_EN = { flights: "Flights", hotels: "Hotels", visa: "Visa", cars: "Transportation" };
  const sectionLabel = (sec) => (accountsLang === "en" ? SECTION_LABELS_EN[sec] : SECTION_LABELS_AR[sec]);

  // Supplier ledger: every supplier that appears on at least one booking, with the
  // total amount owed to them (net price, our cost) minus everything already paid via
  // recorded supplier payments.
  const supplierLedger = (() => {
    const map = {};
    acctBookings.forEach((b) => {
      const name = b.supplier;
      if (!name) return;
      if (!map[name]) map[name] = { supplier: name, sections: new Set(), totalOwed: 0, bookingsCount: 0 };
      map[name].sections.add(b.section);
      map[name].totalOwed += b.net;
      map[name].bookingsCount += 1;
    });
    const paidMap = {};
    supplierPayments.forEach((p) => {
      const name = (p.supplier || "").trim();
      if (!name) return;
      paidMap[name] = (paidMap[name] || 0) + (parseFloat(p.amount) || 0);
    });
    // Suppliers who've only ever received a payment with no booking on file yet still
    // show up, so a stray/advance payment isn't silently dropped from the ledger.
    Object.keys(paidMap).forEach((name) => {
      if (!map[name]) map[name] = { supplier: name, sections: new Set(), totalOwed: 0, bookingsCount: 0 };
    });
    return Object.values(map)
      .map((s) => {
        const paid = paidMap[s.supplier] || 0;
        return { ...s, sections: Array.from(s.sections), paid, balance: s.totalOwed - paid };
      })
      .sort((a, b) => b.balance - a.balance);
  })();

  // Customer ledger: every customer name across all bookings, with the total sold
  // amount owed BY them (split evenly across co-customers on the same booking, the
  // same way per-customer ticket totals are already split elsewhere in the app) minus
  // everything already collected via recorded customer payments.
  const customerLedger = (() => {
    const map = {};
    acctBookings.forEach((b) => {
      const names = b.customers.length ? b.customers : [];
      if (!names.length) return;
      const share = b.sold / names.length;
      names.forEach((name) => {
        if (!map[name]) map[name] = { customer: name, sections: new Set(), totalDue: 0, bookingsCount: 0 };
        map[name].sections.add(b.section);
        map[name].totalDue += share;
        map[name].bookingsCount += 1;
      });
    });
    const paidMap = {};
    customerPayments.forEach((p) => {
      const name = (p.customer || "").trim();
      if (!name) return;
      paidMap[name] = (paidMap[name] || 0) + (parseFloat(p.amount) || 0);
    });
    Object.keys(paidMap).forEach((name) => {
      if (!map[name]) map[name] = { customer: name, sections: new Set(), totalDue: 0, bookingsCount: 0 };
    });
    return Object.values(map)
      .map((c) => {
        const paid = paidMap[c.customer] || 0;
        return { ...c, sections: Array.from(c.sections), paid, balance: c.totalDue - paid };
      })
      .sort((a, b) => b.balance - a.balance);
  })();

  // Current balance of one treasury account: its opening balance, plus every customer
  // receipt and manual "in" entry posted to it, minus every supplier payment, expense,
  // and manual "out" entry posted to it. Everything in the Accounts module is tracked
  // in EGP (matching the EGP-normalized ledgers above), so this is a plain sum.
  const treasuryBalance = (accountId) => {
    const acc = treasuryAccounts.find((a) => a.id === accountId);
    if (!acc) return 0;
    const opening = parseFloat(acc.openingBalance) || 0;
    const inFromCustomers = customerPayments
      .filter((p) => p.accountId === accountId)
      .reduce((s, p) => s + (parseFloat(p.amount) || 0), 0);
    const outToSuppliers = supplierPayments
      .filter((p) => p.accountId === accountId)
      .reduce((s, p) => s + (parseFloat(p.amount) || 0), 0);
    const outExpenses = expenses
      .filter((e) => e.accountId === accountId)
      .reduce((s, e) => s + (parseFloat(e.amount) || 0), 0);
    const manualIn = treasuryEntries
      .filter((e) => e.accountId === accountId && e.direction === "in")
      .reduce((s, e) => s + (parseFloat(e.amount) || 0), 0);
    const manualOut = treasuryEntries
      .filter((e) => e.accountId === accountId && e.direction === "out")
      .reduce((s, e) => s + (parseFloat(e.amount) || 0), 0);
    return opening + inFromCustomers + manualIn - outToSuppliers - outExpenses - manualOut;
  };
  const totalTreasuryBalance = treasuryAccounts.reduce((sum, a) => sum + treasuryBalance(a.id), 0);

  // A single chronological transactions feed for the Treasury tab, merging customer
  // receipts, supplier payments, expenses, and manual entries into one shape.
  const treasuryTransactions = [
    ...customerPayments.map((p) => ({
      id: `cp-${p.id}`, date: p.date, accountId: p.accountId, direction: "in",
      label: `تحصيل من ${p.customer || "-"}`, note: p.note, amount: parseFloat(p.amount) || 0,
    })),
    ...supplierPayments.map((p) => ({
      id: `sp-${p.id}`, date: p.date, accountId: p.accountId, direction: "out",
      label: `دفعة لـ ${p.supplier || "-"}`, note: p.note, amount: parseFloat(p.amount) || 0,
    })),
    ...expenses.map((e) => ({
      id: `ex-${e.id}`, date: e.date, accountId: e.accountId, direction: "out",
      label: e.category + (e.description ? ` - ${e.description}` : ""), note: e.note, amount: parseFloat(e.amount) || 0,
    })),
    ...treasuryEntries.map((e) => ({
      id: `te-${e.id}`, date: e.date, accountId: e.accountId, direction: e.direction,
      label: e.category, note: e.note, amount: parseFloat(e.amount) || 0,
    })),
  ].sort((a, b) => (b.date || "").localeCompare(a.date || ""));

  // Date-range predicate for the Reports tab: "today" | "month" (current calendar
  // month) | "custom" (reportsFrom/reportsTo, either end optional).
  const inReportsRange = (dateStr) => {
    if (!dateStr) return false;
    if (reportsRange === "today") return dateStr === todayDateStr();
    if (reportsRange === "month") return dateStr.slice(0, 7) === todayDateStr().slice(0, 7);
    if (reportsFrom && dateStr < reportsFrom) return false;
    if (reportsTo && dateStr > reportsTo) return false;
    return true;
  };
  const reportRevenueBySection = {
    flights: tickets.filter((t) => inReportsRange(t.date)).reduce((s, t) => s + profitAfterRefund(t), 0),
    hotels: hotelBookings.filter((h) => inReportsRange(h.bookingDate)).reduce((s, h) => s + hotelProfitTotal(h), 0),
    visa: visaBookings
      .filter((v) => inReportsRange(v.bookingDate))
      .reduce((s, v) => s + hotelInEgp(visaProfitTotal(v), v.currency), 0),
    cars: carBookings
      .filter((c) => inReportsRange(c.bookingDate))
      .reduce((s, c) => s + hotelInEgp((parseFloat(c.soldPrice) || 0) - (parseFloat(c.netPrice) || 0), c.currency), 0),
  };
  const reportTotalRevenue = Object.values(reportRevenueBySection).reduce((a, b) => a + b, 0);
  const reportExpensesByCategory = {};
  expenses
    .filter((e) => inReportsRange(e.date))
    .forEach((e) => {
      const cat = e.category || "أخرى";
      reportExpensesByCategory[cat] = (reportExpensesByCategory[cat] || 0) + (parseFloat(e.amount) || 0);
    });
  const reportTotalExpenses = Object.values(reportExpensesByCategory).reduce((a, b) => a + b, 0);
  const reportNetProfit = reportTotalRevenue - reportTotalExpenses;
  const reportBookingsCount = {
    flights: tickets.filter((t) => inReportsRange(t.date)).length,
    hotels: hotelBookings.filter((h) => inReportsRange(h.bookingDate)).length,
    visa: visaBookings.filter((v) => inReportsRange(v.bookingDate)).length,
    cars: carBookings.filter((c) => inReportsRange(c.bookingDate)).length,
  };

  // This month's totals, used on the Overview cards regardless of whatever range is
  // currently selected on the Reports tab.
  const thisMonthPrefix = todayDateStr().slice(0, 7);
  const monthRevenue =
    tickets.filter((t) => (t.date || "").slice(0, 7) === thisMonthPrefix).reduce((s, t) => s + profitAfterRefund(t), 0) +
    hotelBookings.filter((h) => (h.bookingDate || "").slice(0, 7) === thisMonthPrefix).reduce((s, h) => s + hotelProfitTotal(h), 0) +
    visaBookings
      .filter((v) => (v.bookingDate || "").slice(0, 7) === thisMonthPrefix)
      .reduce((s, v) => s + hotelInEgp(visaProfitTotal(v), v.currency), 0) +
    carBookings
      .filter((c) => (c.bookingDate || "").slice(0, 7) === thisMonthPrefix)
      .reduce((s, c) => s + hotelInEgp((parseFloat(c.soldPrice) || 0) - (parseFloat(c.netPrice) || 0), c.currency), 0);
  const monthExpenses = expenses
    .filter((e) => (e.date || "").slice(0, 7) === thisMonthPrefix)
    .reduce((s, e) => s + (parseFloat(e.amount) || 0), 0);
  const totalSupplierBalance = supplierLedger.reduce((s, x) => s + x.balance, 0);
  const totalCustomerBalance = customerLedger.reduce((s, x) => s + x.balance, 0);

  const filteredSupplierLedger = supplierLedger.filter((s) =>
    s.supplier.toLowerCase().includes(supplierQuery.trim().toLowerCase())
  );
  const filteredCustomerLedger = customerLedger.filter((c) =>
    c.customer.toLowerCase().includes(customerQuery.trim().toLowerCase())
  );
  const filteredExpenses = expenses
    .filter((e) => !expenseCategoryFilter || e.category === expenseCategoryFilter)
    .slice()
    .sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  const filteredTreasuryTransactions = treasuryTransactions.filter(
    (tx) => !treasuryFilterAccountId || tx.accountId === treasuryFilterAccountId
  );

  const genId = () => `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

  const handleSaveExpense = () => {
    if (!expenseForm.date || !expenseForm.category || expenseForm.amount === "" || !expenseForm.accountId) {
      setAccountsError("من فضلك أكمل التاريخ والتصنيف والمبلغ والخزينة/الحساب");
      return;
    }
    setAccountsError("");
    const record = { ...expenseForm, id: expenseForm.id || genId() };
    const next = expenseEditingId
      ? expenses.map((e) => (e.id === expenseEditingId ? record : e))
      : [record, ...expenses];
    persistExpenses(next);
    setExpenseForm(getEmptyExpenseForm());
    setExpenseEditingId(null);
    setShowExpenseForm(false);
  };
  const handleEditExpenseClick = (e) => {
    setExpenseForm({ ...e });
    setExpenseEditingId(e.id);
    setShowExpenseForm(true);
  };
  const handleDeleteExpense = (id) => {
    requestConfirm("هل تريد حذف هذا المصروف؟", () => {
      persistExpenses(expenses.filter((e) => e.id !== id));
    });
  };

  const handleSaveTreasuryAccount = () => {
    if (!treasuryForm.name.trim()) {
      setAccountsError("من فضلك اكتب اسم الحساب/الخزينة");
      return;
    }
    setAccountsError("");
    const record = { ...treasuryForm, id: treasuryForm.id || genId() };
    const next = treasuryAccountEditingId
      ? treasuryAccounts.map((a) => (a.id === treasuryAccountEditingId ? record : a))
      : [...treasuryAccounts, record];
    persistTreasuryAccounts(next);
    setTreasuryForm(getEmptyTreasuryAccountForm());
    setTreasuryAccountEditingId(null);
    setShowTreasuryAccountForm(false);
  };
  const handleEditTreasuryAccountClick = (a) => {
    setTreasuryForm({ ...a });
    setTreasuryAccountEditingId(a.id);
    setShowTreasuryAccountForm(true);
  };
  const handleDeleteTreasuryAccount = (id) => {
    requestConfirm("هل تريد حذف هذا الحساب؟ لن يتم حذف الحركات المسجلة عليه من قبل.", () => {
      persistTreasuryAccounts(treasuryAccounts.filter((a) => a.id !== id));
    });
  };

  const handleSaveTreasuryEntry = () => {
    if (!treasuryEntryForm.accountId || treasuryEntryForm.amount === "") {
      setAccountsError("من فضلك اختر الحساب واكتب المبلغ");
      return;
    }
    setAccountsError("");
    const record = { ...treasuryEntryForm, id: genId() };
    persistTreasuryEntries([record, ...treasuryEntries]);
    setTreasuryEntryForm(getEmptyTreasuryEntryForm());
    setShowTreasuryEntryForm(false);
  };
  const handleDeleteTreasuryEntry = (id) => {
    requestConfirm("هل تريد حذف هذا القيد؟", () => {
      persistTreasuryEntries(treasuryEntries.filter((e) => e.id !== id));
    });
  };

  const handleSaveSupplierPayment = () => {
    if (!supplierPaymentForm.supplier || supplierPaymentForm.amount === "" || !supplierPaymentForm.accountId) {
      setAccountsError("من فضلك اختر المورد والخزينة/الحساب واكتب المبلغ");
      return;
    }
    setAccountsError("");
    const record = { ...supplierPaymentForm, id: genId() };
    persistSupplierPayments([record, ...supplierPayments]);
    setSupplierPaymentForm({ ...getEmptySupplierPaymentForm(), supplier: supplierPaymentForm.supplier });
  };
  const handleDeleteSupplierPayment = (id) => {
    requestConfirm("هل تريد حذف هذه الدفعة؟", () => {
      persistSupplierPayments(supplierPayments.filter((p) => p.id !== id));
    });
  };

  const handleSaveCustomerPayment = () => {
    if (!customerPaymentForm.customer || customerPaymentForm.amount === "" || !customerPaymentForm.accountId) {
      setAccountsError("من فضلك اختر العميل والخزينة/الحساب واكتب المبلغ");
      return;
    }
    setAccountsError("");
    const record = { ...customerPaymentForm, id: genId() };
    persistCustomerPayments([record, ...customerPayments]);
    setCustomerPaymentForm({ ...getEmptyCustomerPaymentForm(), customer: customerPaymentForm.customer });
  };
  const handleDeleteCustomerPayment = (id) => {
    requestConfirm("هل تريد حذف هذا التحصيل؟", () => {
      persistCustomerPayments(customerPayments.filter((p) => p.id !== id));
    });
  };

  // Exports the currently selected report range to an Excel workbook — revenue by
  // section, expenses by category, and the net profit summary — following the same
  // XLSX.utils/writeFile pattern used by the other sections' exports in this file.
  const handleExportAccountsReport = () => {
    const wb = XLSX.utils.book_new();
    const summaryRows = [
      { "البند": "إيرادات الطيران", "المبلغ (ج.م)": Math.round(reportRevenueBySection.flights * 100) / 100 },
      { "البند": "إيرادات الفنادق", "المبلغ (ج.م)": Math.round(reportRevenueBySection.hotels * 100) / 100 },
      { "البند": "إيرادات الفيزا", "المبلغ (ج.م)": Math.round(reportRevenueBySection.visa * 100) / 100 },
      { "البند": "إيرادات الترانسفير", "المبلغ (ج.م)": Math.round(reportRevenueBySection.cars * 100) / 100 },
      { "البند": "إجمالي الإيرادات", "المبلغ (ج.م)": Math.round(reportTotalRevenue * 100) / 100 },
      { "البند": "إجمالي المصروفات", "المبلغ (ج.م)": Math.round(reportTotalExpenses * 100) / 100 },
      { "البند": "صافي الربح", "المبلغ (ج.م)": Math.round(reportNetProfit * 100) / 100 },
    ];
    const summarySheet = XLSX.utils.json_to_sheet(summaryRows);
    XLSX.utils.book_append_sheet(wb, summarySheet, "Summary");
    const expenseRows = Object.entries(reportExpensesByCategory).map(([cat, amt]) => ({
      "التصنيف": cat,
      "المبلغ (ج.م)": Math.round(amt * 100) / 100,
    }));
    if (expenseRows.length) {
      const expenseSheet = XLSX.utils.json_to_sheet(expenseRows);
      XLSX.utils.book_append_sheet(wb, expenseSheet, "Expenses");
    }
    const supplierRows = supplierLedger.map((s) => ({
      "المورد": s.supplier,
      "إجمالى المستحق": Math.round(s.totalOwed * 100) / 100,
      "المدفوع": Math.round(s.paid * 100) / 100,
      "المتبقى": Math.round(s.balance * 100) / 100,
    }));
    if (supplierRows.length) {
      const supplierSheet = XLSX.utils.json_to_sheet(supplierRows);
      XLSX.utils.book_append_sheet(wb, supplierSheet, "Suppliers");
    }
    const customerRows = customerLedger.map((c) => ({
      "العميل": c.customer,
      "إجمالى المستحق": Math.round(c.totalDue * 100) / 100,
      "المحصل": Math.round(c.paid * 100) / 100,
      "المتبقى": Math.round(c.balance * 100) / 100,
    }));
    if (customerRows.length) {
      const customerSheet = XLSX.utils.json_to_sheet(customerRows);
      XLSX.utils.book_append_sheet(wb, customerSheet, "Customers");
    }
    XLSX.writeFile(wb, `accounts_report_${todayDateStr()}.xlsx`);
  };

  // ---------- Files ----------
  const visibleFiles = (
    !currentUser
      ? []
      : filesPerm.canViewAll
      ? files
      : files.filter((f) =>
          f.employeeUsername ? f.employeeUsername === currentUser.username : f.createdBy === currentUser.name
        )
  )
    // Ordered by the file's own date (newest first), with the serial as a tie-breaker
    // for same-day files — the list always follows the dates rather than raw creation/
    // array order.
    .slice()
    .sort((a, b) => (b.createdAt || "").localeCompare(a.createdAt || "") || (b.serial || "").localeCompare(a.serial || ""));

  const FILE_SOURCE_LABELS = { flights: "Flight", hotels: "Hotel", visa: "Visa", cars: "Transportation" };


  // ---------- Files: search + filters ----------
  const fileYearsAvailable = Array.from(
    new Set(visibleFiles.map((f) => (f.createdAt ? f.createdAt.slice(0, 4) : "")).filter(Boolean))
  ).sort((a, b) => b.localeCompare(a));
  const fileCompaniesAvailable = Array.from(
    new Set(visibleFiles.map((f) => (f.company || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));
  const fileEmployeesAvailable = Array.from(
    new Set(visibleFiles.map((f) => (f.createdBy || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));

  const hasActiveFileFilter = !!(fileSelectedYear.length || fileSelectedCompany.length || fileSelectedEmployee.length || fileQuery.trim());
  const activeFileFilterCount = fileSelectedYear.length + fileSelectedCompany.length + fileSelectedEmployee.length + (fileQuery.trim() ? 1 : 0);
  const clearAllFileFilters = () => {
    setFileQuery("");
    setFileSelectedYear([]);
    setFileSelectedCompany([]);
    setFileSelectedEmployee([]);
  };
  const filteredFiles = visibleFiles.filter((f) => {
    if (fileSelectedYear.length && !fileSelectedYear.includes((f.createdAt || "").slice(0, 4))) return false;
    if (fileSelectedCompany.length && !fileSelectedCompany.includes((f.company || "").trim())) return false;
    if (fileSelectedEmployee.length && !fileSelectedEmployee.includes((f.createdBy || "").trim())) return false;
    const q = fileQuery.trim().toLowerCase();
    if (!q) return true;
    return (
      (f.serial || "").toLowerCase().includes(q) ||
      (f.company || "").toLowerCase().includes(q) ||
      (f.notes || "").toLowerCase().includes(q) ||
      (f.createdBy || "").toLowerCase().includes(q)
    );
  });

  // Auto-generates a starting serial number for a file, based on the file's own date
  // (defaults to today, but the date is user-editable — see updateFileDate below):
  // F-YYYYMMDD-00001, F-YYYYMMDD-00002, ... The trailing 5-digit running number is
  // GLOBAL across every file ever created (not per-date): it always continues from
  // one more than the highest running number found anywhere in the file list, so it
  // keeps climbing steadily no matter what date a file is given. This is only a
  // suggested starting value: the serial is a plain text field the user can freely
  // retype per file afterwards (see the "Serial" input on the open file panel), so
  // it isn't locked to this pattern.
  // Computed off the full (unfiltered) files list so numbering stays globally consistent
  // no matter who's creating/editing the file.
  const nextFileSerial = (list, dateStr) => {
    const datePart = (dateStr || todayDateStr()).replace(/-/g, "");
    const prefix = `F-${datePart}-`;
    const maxN = (list || []).reduce((max, f) => {
      const match = (f.serial || "").match(/(\d{5})$/); // last 5 digits, wherever in the serial
      const n = match ? parseInt(match[1], 10) : 0;
      return Math.max(max, n);
    }, 0);
    return `${prefix}${String(maxN + 1).padStart(5, "0")}`;
  };

  // Builds a read-only price snapshot of a ticket/hotel/visa record to drop into a
  // file. This is a COPY only — it never references or mutates the original record, so
  // adding it to a file has no effect whatsoever on the Flights/Hotels/Visa totals.
  const buildFileItem = (sourceType, record) => {
    const base = { id: `FI-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`, sourceType, sourceId: record.id };
    if (sourceType === "flights") {
      const names = getCustomers(record).map((c) => c.name).filter(Boolean).join(", ");
      return {
        ...base,
        label: `${routeLabel(record)}${names ? " · " + names : ""}`,
        date: record.date,
        currency: "EGP",
        netPrice: parseFloat(record.netPrice) || 0,
        soldPrice: parseFloat(record.soldPrice) || 0,
      };
    }
    if (sourceType === "hotels") {
      return {
        ...base,
        label: `${record.hotel || "Hotel"}${record.customer ? " · " + record.customer : ""}`,
        date: record.bookingDate,
        currency: "EGP",
        netPrice: hotelNetTotal(record),
        soldPrice: hotelSoldTotal(record),
      };
    }
    if (sourceType === "visa") {
      const names = (record.customers || []).map((c) => c.name).filter(Boolean).join(", ");
      return {
        ...base,
        label: `${record.visaType || "Visa"}${names ? " · " + names : ""}`,
        date: record.bookingDate,
        currency: record.currency || "EGP",
        netPrice: visaNetTotal(record),
        soldPrice: visaSoldTotal(record),
      };
    }
    return { ...base, label: "-", date: "", currency: "EGP", netPrice: 0, soldPrice: 0 };
  };

  // Every item's amount converted into EGP (same conversion hotels already use), so a
  // file mixing EGP and USD items still totals correctly.
  const fileTotals = (f) =>
    (f.items || []).reduce(
      (acc, it) => {
        acc.net += hotelInEgp(it.netPrice, it.currency);
        acc.sold += hotelInEgp(it.soldPrice, it.currency);
        acc.profit = acc.sold - acc.net;
        return acc;
      },
      { net: 0, sold: 0, profit: 0 }
    );

  const filesGrandTotals = filteredFiles.reduce(
    (acc, f) => {
      const t = fileTotals(f);
      acc.net += t.net;
      acc.sold += t.sold;
      acc.profit += t.profit;
      return acc;
    },
    { net: 0, sold: 0, profit: 0 }
  );

  const updateFileField = async (id, field, value) => {
    await persistFiles(files.map((f) => (f.id === id ? { ...f, [field]: value } : f)));
  };

  // Lets the user control a file's date directly (it isn't locked to the day the file was
  // created). Changing the date also re-generates the serial: the date part is updated to
  // match, and the trailing running number is recomputed as one more than the current
  // global maximum across every other file (so it keeps climbing steadily regardless of
  // date), and the time part reflects the moment the change was made.
  const updateFileDate = async (id, newDate) => {
    const others = files.filter((f) => f.id !== id);
    const newSerial = nextFileSerial(others, newDate);
    await persistFiles(
      files.map((f) => (f.id === id ? { ...f, createdAt: newDate, serial: newSerial } : f))
    );
  };

  const addItemToFile = async (fileId, sourceType, record) => {
    if (!currentUser) return;
    const item = buildFileItem(sourceType, record);
    await persistFiles(files.map((f) => (f.id === fileId ? { ...f, items: [...(f.items || []), item] } : f)));
  };

  const removeItemFromFile = async (fileId, itemId) => {
    if (!canEditTickets) return;
    await persistFiles(
      files.map((f) => (f.id === fileId ? { ...f, items: (f.items || []).filter((i) => i.id !== itemId) } : f))
    );
  };

  // Unlike deleting elsewhere in the app, deleting a FILE is intentionally open to every
  // signed-in employee (not gated by canDeleteTickets) — same reasoning as adding items to
  // a file above: files are a shared working space, not permission-gated per employee.
  const deleteFile = async (id) => {
    if (!currentUser) return;
    await persistFiles(files.filter((f) => f.id !== id));
    if (openFileId === id) setOpenFileId(null);
  };

  const openFile = openFileId ? files.find((f) => f.id === openFileId) : null;

  // Used by the "copy to a file" button on the Flights/Hotels/Visa tables: drops a
  // snapshot of that one record into the chosen file, without touching the record itself.
  const copySourceToFile = async (fileId) => {
    if (!copyPickerSource) return;
    await addItemToFile(fileId, copyPickerSource.type, copyPickerSource.record);
    setCopyPickerSource(null);
  };

  // "New file" shortcut inside the copy picker: creates the file, then immediately
  // drops the pending copy into it.
  const createFileAndCopySource = async () => {
    if (!currentUser || !copyPickerSource) return;
    const record = {
      id: `FL-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      serial: nextFileSerial(files),
      createdAt: todayDateStr(),
      createdBy: currentUser.name,
      employeeUsername: currentUser.username,
      company: "",
      notes: "",
      items: [buildFileItem(copyPickerSource.type, copyPickerSource.record)],
    };
    await persistFiles([record, ...files]);
    setCopyPickerSource(null);
  };

  // Starts a new file in "draft" mode: nothing is saved to the files table yet, but the
  // serial number is generated right away (based on today's date) so it's visible while
  // filling in the rest. The user fills in details and can pull in service copies, then
  // presses "Add file" to confirm.
  const startNewFileDraft = () => {
    if (!currentUser) return;
    const createdAt = todayDateStr();
    setDraftFile({ serial: nextFileSerial(files, createdAt), company: "", notes: "", createdAt, items: [] });
  };

  const updateDraftField = (field, value) =>
    setDraftFile((d) => (d ? { ...d, [field]: value } : d));

  // Changing the draft's date re-generates its serial to match (same as updateFileDate
  // does for an already-saved file), so the serial shown always reflects the file's date.
  const updateDraftDate = (newDate) =>
    setDraftFile((d) => (d ? { ...d, createdAt: newDate, serial: nextFileSerial(files, newDate) } : d));

  const addDraftItem = (sourceType, record) =>
    setDraftFile((d) => (d ? { ...d, items: [...(d.items || []), buildFileItem(sourceType, record)] } : d));

  const removeDraftItem = (itemId) =>
    setDraftFile((d) => (d ? { ...d, items: (d.items || []).filter((i) => i.id !== itemId) } : d));

  const cancelDraftFile = () => setDraftFile(null);

  // "Add file" (confirm) button on the draft panel: this is the moment the file actually
  // gets created and placed in the main files table, using the serial that was already
  // generated (and shown) while in draft mode.
  const confirmDraftFile = async () => {
    if (!currentUser || !draftFile) return;
    const record = {
      id: `FL-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      serial: draftFile.serial || nextFileSerial(files, draftFile.createdAt),
      createdAt: draftFile.createdAt || todayDateStr(),
      createdBy: currentUser.name,
      employeeUsername: currentUser.username,
      company: draftFile.company || "",
      notes: draftFile.notes || "",
      items: draftFile.items || [],
    };
    await persistFiles([record, ...files]);
    setDraftFile(null);
  };

  const monthsAvailable = Array.from(new Set(visibleTickets.map((t) => monthKey(t.date)))).sort((a, b) =>
    b.localeCompare(a)
  );

  const yearsAvailable = Array.from(
    new Set(
      visibleTickets
        .map((t) => (t.date ? t.date.slice(0, 4) : ""))
        .filter(Boolean)
    )
  ).sort((a, b) => b.localeCompare(a));

  const companiesAvailable = Array.from(
    new Set(visibleTickets.map((t) => (t.company || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));

  const employeesAvailable = Array.from(
    new Set(visibleTickets.map((t) => (t.employee || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));

  const suppliersAvailable = Array.from(
    new Set(visibleTickets.map((t) => (t.supplier || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));

  const byMonth = selectedMonth.length
    ? visibleTickets.filter((t) => selectedMonth.includes(monthKey(t.date)))
    : visibleTickets;

  const byYear = selectedYear.length
    ? byMonth.filter((t) => selectedYear.includes((t.date || "").slice(0, 4)))
    : byMonth;

  const byCompany = selectedCompany.length
    ? byYear.filter((t) => selectedCompany.includes((t.company || "").trim()))
    : byYear;

  const byEmployee = selectedEmployee.length
    ? byCompany.filter((t) => selectedEmployee.includes((t.employee || "").trim()))
    : byCompany;

  const bySupplier = selectedSupplier.length
    ? byEmployee.filter((t) => selectedSupplier.includes((t.supplier || "").trim()))
    : byEmployee;

  const filtered = bySupplier.filter((t) => {
    const q = query.trim().toLowerCase();
    if (!q) return true;
    const customers = getCustomers(t);
    return (
      (t.employee || "").toLowerCase().includes(q) ||
      (t.company || "").toLowerCase().includes(q) ||
      (t.from || "").toLowerCase().includes(q) ||
      (t.to || "").toLowerCase().includes(q) ||
      (Array.isArray(t.destinations) ? t.destinations.join(" ") : "").toLowerCase().includes(q) ||
      (t.airline || "").toLowerCase().includes(q) ||
      customers.some(
        (c) =>
          (c.name || "").toLowerCase().includes(q) ||
          (c.ticketNumber || "").toLowerCase().includes(q) ||
          (c.pnrReference || "").toLowerCase().includes(q)
      )
    );
  });

  // Sort tickets by issue date (most recent first). Tickets with no date are pushed
  // to the end instead of being sorted arbitrarily.
  const sortedFiltered = [...filtered].sort((a, b) => {
    if (!a.date && !b.date) return 0;
    if (!a.date) return 1;
    if (!b.date) return -1;
    return b.date.localeCompare(a.date);
  });

  // Ticket numbers that have been superseded by a reissue — i.e. every ticket number
  // that appears as some OTHER ticket's oldTicketNumber. These are hidden from the main
  // table by default (the new, reissued ticket shows in its place, at its own date), so
  // an old ticket number doesn't keep cluttering the list and the date order never gets
  // reshuffled to nest a reissued ticket under an older row. Clicking that new ticket's
  // "Exchanged" badge (see jumpToRow above) reveals the specific old row again.
  const supersededTicketNumbers = new Set();
  tickets.forEach((t2) => {
    if (t2.isReissued && t2.oldTicketNumber) {
      supersededTicketNumbers.add(t2.oldTicketNumber.trim().toUpperCase());
    }
  });

  // The ticket currently open in the detail "page", if any.
  const viewingTicket = viewingTicketId ? visibleTickets.find((t) => t.id === viewingTicketId) : null;

  // Counts and sums per CUSTOMER rather than per ticket/booking: a booking with several
  // customers contributes its full (unsplit) total/profit once for each customer, and
  // each customer counts as one ticket. This keeps the summary cards, monthly totals,
  // and company breakdown consistent with the per-customer rows shown in the ticket table.
  // A recorded refund is a real, one-time amount for the booking, so it's deducted once
  // (not multiplied by customer count) — reducing sales by what went back to the customer
  // and adjusting profit by that same amount net of whatever the airline refunded back.
  const countAndSum = (rows) =>
    rows.reduce(
      (acc, t) => {
        const n = getCustomers(t).length || 1;
        const refundCustomerAmt = getRefunds(t).reduce((s, r) => s + (parseFloat(r.customerAmount) || 0), 0);
        const refundAirlineAmt = getRefunds(t).reduce((s, r) => s + (parseFloat(r.airlineAmount) || 0), 0);
        acc.count += n;
        acc.total += (parseFloat(t.soldPrice) || 0) * n - refundCustomerAmt;
        acc.profit += profit(t.netPrice, t.soldPrice) * n + refundAirlineAmt - refundCustomerAmt;
        return acc;
      },
      { count: 0, total: 0, profit: 0 }
    );

  const totals = countAndSum(bySupplier);

  const monthlyBreakdown = monthsAvailable.map((key) => {
    const rows = visibleTickets.filter((t) => monthKey(t.date) === key);
    return { key, ...countAndSum(rows) };
  });

  const companyBreakdown = companiesAvailable.map((name) => {
    const rows = visibleTickets.filter((t) => (t.company || "").trim() === name);
    const customers = Array.from(
      new Set(rows.flatMap((t) => getCustomers(t).map((c) => c.name)).filter(Boolean))
    ).sort((a, b) => a.localeCompare(b));
    return { name, customers, ...countAndSum(rows) };
  });

  // Ticket-level status text used in the exported "Status" column, replacing the old
  // per-customer numbering. "Reissued" applies to the whole booking (shown on the first
  // customer's row); "Refunded" applies only to the specific customer the refund was
  // recorded against, since a multi-customer booking may have just one refunded ticket.
  const ticketStatus = (t, i) => {
    const parts = [];
    if (i === 0 && t.isReissued) parts.push("Exchanged");
    if (refundForIndex(t, i)) parts.push("Refunded");
    return parts.join(" & ");
  };

  // Builds the per-customer row list for one ticket set, sorted by issue date
  // (earliest first; undated tickets pushed to the end). Tickets issued on the
  // SAME date are then ordered by ticket number ascending (numeric-aware, so
  // "077-1234567890" sorts before "077-1234567900" correctly).
  const ticketRows = (rows) => {
    const firstTicketNumber = (t) => (getCustomers(t)[0] && getCustomers(t)[0].ticketNumber) || "";
    const sorted = [...rows].sort((a, b) => {
      if (!a.date && !b.date) {
        return firstTicketNumber(a).localeCompare(firstTicketNumber(b), undefined, { numeric: true, sensitivity: "base" });
      }
      if (!a.date) return 1;
      if (!b.date) return -1;
      const dateCompare = a.date.localeCompare(b.date);
      if (dateCompare !== 0) return dateCompare;
      return firstTicketNumber(a).localeCompare(firstTicketNumber(b), undefined, { numeric: true, sensitivity: "base" });
    });
    return sorted.flatMap((t) => {
      const customers = getCustomers(t);
      const rows = customers.map((c, i) => ({
        "Type": "Ticket",
        "Employee": t.employee || "",
        "Company": t.company || "",
        "Supplier": t.supplier || "",
        "Status": ticketStatus(t, i),
        "Customer": c.name || "",
        "Ticket number": c.ticketNumber || "",
        "PNR reference": c.pnrReference || "",
        "From": t.from,
        "To": t.to,
        "Route": routeLabel(t),
        "Airline": t.airline || "",
        "Issue date": t.date ? formatDisplayDate(t.date) : "",
        // Net/sold price and profit are the ORIGINAL booking amounts, shown once on the
        // first customer's row. A recorded refund gets its own row directly underneath
        // (see below) instead of adding more columns here — keeps this row short and
        // keeps the refund's figures clearly separate from the original price.
        "Net price": i === 0 ? parseFloat(t.netPrice) || 0 : "",
        "Sold price": i === 0 ? parseFloat(t.soldPrice) || 0 : "",
        "Profit": i === 0 ? profit(t.netPrice, t.soldPrice) : "",
        "Refund date": "",
        "Refund (airline)": "",
        "Refund (customer)": "",
        "Net after refund": "",
        "Sold after refund": "",
        "Profit after refund": "",
        "Notes": t.notes || "",
      }));
      // Every recorded refund gets its own short row directly under the original
      // ticket's row(s) — just that refund's figures, so neither row is cluttered with
      // columns that don't apply to it. A booking with several refunded customers gets
      // one refund row per customer; the running after-refund totals (which reflect all
      // of that booking's refunds combined) are only shown once, on the last of them.
      const bookingRefunds = getRefunds(t).filter((r) => r && (r.airlineAmount !== "" || r.customerAmount !== ""));
      bookingRefunds.forEach((refund, ri) => {
        const refundedCustomer = customers[refund.customerIndex || 0] || customers[0];
        const isLast = ri === bookingRefunds.length - 1;
        rows.push({
          "Type": "Refund",
          "Employee": t.employee || "",
          "Company": t.company || "",
          "Supplier": "",
          "Status": "Refunded",
          "Customer": refundedCustomer ? refundedCustomer.name || "" : "",
          "Ticket number": `Refund — ${(refundedCustomer && refundedCustomer.ticketNumber) || firstTicketNumber(t) || "ticket"}`,
          "From": "",
          "To": "",
          "Route": "",
          "Airline": "",
          "Issue date": "",
          "Net price": "",
          "Sold price": "",
          "Profit": "",
          "Refund date": refund.date ? formatDisplayDate(refund.date) : "",
          "Refund (airline)": parseFloat(refund.airlineAmount) || 0,
          "Refund (customer)": parseFloat(refund.customerAmount) || 0,
          "Net after refund": isLast ? netAfterRefund(t) : "",
          "Sold after refund": isLast ? soldAfterRefund(t) : "",
          "Profit after refund": isLast ? profitAfterRefund(t) : "",
          "Notes": "",
        });
      });
      return rows;
    });
  };

  // Sums the original price columns, the refund columns, and the resulting after-refund
  // figures across a raw ticket list (once per booking, matching how those columns are
  // only populated on each booking's first/refunded row above).
  const sumTicketPrices = (rows) =>
    rows.reduce(
      (acc, t) => {
        acc.net += parseFloat(t.netPrice) || 0;
        acc.sold += parseFloat(t.soldPrice) || 0;
        acc.profit += profit(t.netPrice, t.soldPrice);
        acc.refundAirline += getRefunds(t).reduce((s, r) => s + (parseFloat(r.airlineAmount) || 0), 0);
        acc.refundCustomer += getRefunds(t).reduce((s, r) => s + (parseFloat(r.customerAmount) || 0), 0);
        acc.netAfter += netAfterRefund(t);
        acc.soldAfter += soldAfterRefund(t);
        acc.profitAfter += profitAfterRefund(t);
        return acc;
      },
      { net: 0, sold: 0, profit: 0, refundAirline: 0, refundCustomer: 0, netAfter: 0, soldAfter: 0, profitAfter: 0 }
    );

  // Appends a totals row (original + refund + after-refund figures) to the end of a sheet's rows.
  const rowsWithTotals = (rows) => {
    const sums = sumTicketPrices(rows);
    return [
      ...ticketRows(rows),
      {
        "Employee": "", "Company": "", "Supplier": "", "Status": "", "Customer": "",
        "Ticket number": "", "From": "", "To": "", "Airline": "", "Issue date": "TOTAL",
        "Net price": Math.round(sums.net * 100) / 100,
        "Sold price": Math.round(sums.sold * 100) / 100,
        "Profit": Math.round(sums.profit * 100) / 100,
        "Refund date": "",
        "Refund (airline)": Math.round(sums.refundAirline * 100) / 100,
        "Refund (customer)": Math.round(sums.refundCustomer * 100) / 100,
        "Net after refund": Math.round(sums.netAfter * 100) / 100,
        "Sold after refund": Math.round(sums.soldAfter * 100) / 100,
        "Profit after refund": Math.round(sums.profitAfter * 100) / 100,
        "Notes": "",
      },
    ];
  };

  // Builds a single, human-readable line describing every filter currently applied
  // (month/year/company/employee/supplier/search), so it can be dropped into one cell
  // at the top of an export instead of forcing whoever opens the file to guess what
  // selection it represents.
  const describeFilters = ({ month, year, company, employee, supplier, search } = {}) => {
    // Each of month/year/company/employee/supplier may be passed as a single value
    // (a plain string, e.g. from exportMonth) or an array of several selected values
    // (from the multi-select filters) — join arrays with an Arabic-style comma.
    const joinVal = (v) => (Array.isArray(v) ? v.join("، ") : v);
    const hasVal = (v) => (Array.isArray(v) ? v.length > 0 : !!v);
    const parts = [];
    if (hasVal(year)) parts.push(`السنة: ${joinVal(year)}`);
    if (hasVal(month)) parts.push(`الشهر: ${Array.isArray(month) ? month.map(monthLabel).join("، ") : monthLabel(month)}`);
    if (hasVal(company)) parts.push(`الشركة: ${joinVal(company)}`);
    if (hasVal(employee)) parts.push(`الموظف: ${joinVal(employee)}`);
    if (hasVal(supplier)) parts.push(`المورد: ${joinVal(supplier)}`);
    if (search) parts.push(`بحث: ${search}`);
    return parts.length ? `الفلاتر المطبقة — ${parts.join("  |  ")}` : "بدون فلاتر — كل النتائج";
  };

  // Writes rows to a sheet with a filter-summary banner merged across one cell at the
  // very top (row 1), a blank spacer row, then the normal header + data rows below.
  const sheetWithFilterBanner = (rows, filterLabel) => {
    const ws = XLSX.utils.json_to_sheet(rows, { origin: "A3" });
    const colCount = Math.max(Object.keys(rows[0] || {}).length, 1);
    XLSX.utils.sheet_add_aoa(ws, [[filterLabel]], { origin: "A1" });
    ws["!merges"] = [
      ...(ws["!merges"] || []),
      { s: { r: 0, c: 0 }, e: { r: 0, c: colCount - 1 } },
    ];
    return ws;
  };

  const exportMonth = (key) => {
    const rows = visibleTickets.filter((t) => monthKey(t.date) === key);
    const ws = sheetWithFilterBanner(rowsWithTotals(rows), describeFilters({ month: key }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Details");
    XLSX.writeFile(wb, `tickets_${key}.xlsx`);
  };

  const exportAllMonths = () => {
    const wb = XLSX.utils.book_new();
    const summarySheet = XLSX.utils.json_to_sheet(
      monthlyBreakdown.map((m) => ({
        "Month": monthLabel(m.key),
        "Tickets": m.count,
        "Total sales": Math.round(m.total * 100) / 100,
        "Total profit": Math.round(m.profit * 100) / 100,
      }))
    );
    XLSX.utils.book_append_sheet(wb, summarySheet, "Monthly totals");

    monthlyBreakdown.forEach((m) => {
      const rows = visibleTickets.filter((t) => monthKey(t.date) === m.key);
      const ws = sheetWithFilterBanner(rowsWithTotals(rows), describeFilters({ month: m.key }));
      const safeName = m.key.replace(/[:\\\/\?\*\[\]]/g, "-").slice(0, 31);
      XLSX.utils.book_append_sheet(wb, ws, safeName);
    });

    XLSX.writeFile(wb, "monthly_ticket_totals.xlsx");
  };

  // Exports exactly the tickets matching the currently selected month / year / company /
  // employee / supplier filters AND the search box (any combination) — the same set of
  // tickets currently shown on screen — sorted by issue date (same-day tickets ordered
  // by ticket number ascending), as a single sheet ending with a totals row.
  const hasActiveFilter = !!(selectedMonth.length || selectedYear.length || selectedCompany.length || selectedEmployee.length || selectedSupplier.length || query.trim());

  // Count of active filters/search, shown as a badge on the "Filters" toggle button so
  // the person can see at a glance how many are applied without opening the panel.
  const activeFilterCount =
    selectedYear.length + selectedMonth.length + selectedCompany.length + selectedEmployee.length + selectedSupplier.length + (query.trim() ? 1 : 0);

  // Resets every filter and the search box at once — used by the "Clear all" action
  // in the filter chips row.
  const clearAllFilters = () => {
    setSelectedYear([]);
    setSelectedMonth([]);
    setSelectedCompany([]);
    setSelectedEmployee([]);
    setSelectedSupplier([]);
    setQuery("");
  };

  const exportFiltered = () => {
    const filterLabel = describeFilters({
      month: selectedMonth,
      year: selectedYear,
      company: selectedCompany,
      employee: selectedEmployee,
      supplier: selectedSupplier,
      search: query.trim(),
    });
    const ws = sheetWithFilterBanner(rowsWithTotals(filtered), filterLabel);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Details");
    const parts = [
      selectedYear,
      selectedMonth,
      selectedCompany,
      selectedEmployee,
      selectedSupplier,
    ]
      .flat()
      .filter(Boolean)
      .map((p) => p.replace(/[^a-zA-Z0-9-]+/g, "_"));
    XLSX.writeFile(wb, `tickets_${parts.length ? parts.join("_") : "filtered"}.xlsx`);
  };

  const fmt = (n) => new Intl.NumberFormat("en-US", { maximumFractionDigits: 2 }).format(n);

  // ---------- Render: loading ----------
  if (loading || setupComplete === null) {
    return (
      <div className="w-full min-h-screen bg-gradient-to-br from-teal-50 via-stone-50 to-white flex items-center justify-center">
        <style>{`@keyframes pdmFadeIn{from{opacity:0}to{opacity:1}} @keyframes pdmFly{0%{transform:translateX(-6px) rotate(45deg)}50%{transform:translateX(6px) rotate(45deg)}100%{transform:translateX(-6px) rotate(45deg)}}`}</style>
        <p className="text-teal-800/60 text-sm flex items-center gap-2" style={{ animation: "pdmFadeIn .3s ease-out both" }}>
          <Plane size={16} className="rotate-45" style={{ animation: "pdmFly 1.4s ease-in-out infinite" }} /> Loading...
        </p>
      </div>
    );
  }

  // ---------- Render: first-run setup (only ever shown once, before any account exists) ----------
  if (employees && employees.length === 0 && !setupComplete) {
    return (
      <div className="w-full min-h-screen bg-gradient-to-br from-teal-50 via-stone-50 to-white flex items-center justify-center p-4">
        <style>{`@keyframes pdmPopIn{from{opacity:0;transform:translateY(10px) scale(0.96)}to{opacity:1;transform:translateY(0) scale(1)}}`}</style>
        <div className="bg-white rounded-2xl border border-stone-200 p-6 w-full max-w-sm shadow-xl shadow-teal-900/5" style={{ animation: "pdmPopIn .3s cubic-bezier(0.16,1,0.3,1) both" }}>
          <div className="flex items-center gap-2 mb-1">
            <div className="bg-teal-800/10 text-teal-800 rounded-xl p-1.5">
              <Lock size={16} />
            </div>
            <h1 className="font-bold text-stone-900">Create the admin account</h1>
          </div>
          <p className="text-xs text-stone-500 mb-4">
            No employees exist yet. Create the first account — it will be the main account, and only it will be able to add or remove other employees.
          </p>
          {loginError && <div className="bg-red-50 text-red-700 text-sm rounded-xl px-3 py-2 mb-3">{loginError}</div>}
          <div className="space-y-3">
            <div>
              <label className="text-xs text-stone-500 block mb-1">Full name</label>
              <input className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={setupName} onChange={(e) => setSetupName(e.target.value)} placeholder="e.g. Sara Ahmed" />
            </div>
            <div>
              <label className="text-xs text-stone-500 block mb-1">Username</label>
              <input className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={setupUsername} onChange={(e) => setSetupUsername(e.target.value)} placeholder="sara" />
            </div>
            <div>
              <label className="text-xs text-stone-500 block mb-1">Password</label>
              <input type="password" className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={setupPassword} onChange={(e) => setSetupPassword(e.target.value)} placeholder="••••••" />
            </div>
          </div>
          <button onClick={handleCreateFirstAdmin}
            className="w-full mt-4 bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 shadow-sm shadow-teal-800/30 ring-1 ring-inset ring-white/10 transition-colors">
            Create account and continue
          </button>
          <p className="text-xs text-stone-400 mt-4">
            Note: this is a simple access gate stored with the app's data, not a secure authentication system — anyone with technical access to the app's data can read stored passwords. Don't reuse an important password here.
          </p>
        </div>
      </div>
    );
  }

  // ---------- Render: accounts missing after setup was already completed ----------
  // Setup has already happened once before, but no employee accounts exist right now
  // (e.g. all accounts were removed, or a restore emptied them). We deliberately do NOT
  // fall back to the unauthenticated first-run setup screen here, since that would let
  // anyone create a brand-new admin account without any credentials.
  if (employees && employees.length === 0 && setupComplete) {
    return (
      <div className="w-full min-h-screen bg-gradient-to-br from-teal-50 via-stone-50 to-white flex items-center justify-center p-4">
        <div className="bg-white rounded-2xl border border-stone-200 p-6 w-full max-w-sm text-center shadow-xl shadow-teal-900/5">
          <Lock size={22} className="text-stone-400 mx-auto mb-2" />
          <h1 className="font-bold text-stone-900 mb-1">No accounts available</h1>
          <p className="text-xs text-stone-500">
            This app was already set up before, but no employee accounts currently exist. Restore a backup that contains employee accounts, or contact whoever manages this app.
          </p>
        </div>
      </div>
    );
  }

  // ---------- Render: login screen ----------
  if (!currentUser) {
    return (
      <div className="w-full min-h-screen relative flex items-center justify-center p-4 overflow-hidden bg-gradient-to-br from-teal-900 via-teal-800 to-[#0d3b3e]" style={{ fontFamily: "'Inter', sans-serif" }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:ital,opsz,wght@0,9..144,500;0,9..144,600;1,9..144,500&family=Inter:wght@400;500;600;700&display=swap');
          @keyframes pdmFadeIn { from { opacity: 0; } to { opacity: 1; } }
          @keyframes pdmPopIn { from { opacity: 0; transform: translateY(18px) scale(0.96); } to { opacity: 1; transform: translateY(0) scale(1); } }
          @keyframes pdmFlyAcross { from { opacity: 0; transform: translateX(-40px) rotate(45deg); } to { opacity: 0.7; transform: translateX(0) rotate(45deg); } }
          @media (prefers-reduced-motion: no-preference) {
            button:not(:disabled) { transition: transform .15s ease, box-shadow .15s ease, background-color .15s ease !important; }
            button:not(:disabled):active { transform: scale(0.97); }
            input { transition: box-shadow .15s ease, border-color .15s ease; }
          }
        `}</style>
        {/* Decorative sky + route backdrop */}
        <div className="pointer-events-none absolute inset-0 overflow-hidden">
          <div
            className="absolute inset-0 opacity-[0.15]"
            style={{ backgroundImage: "radial-gradient(circle, #ffffff 1px, transparent 1px)", backgroundSize: "26px 26px" }}
          />
          <div className="absolute -top-24 -left-24 w-72 h-72 bg-teal-400/25 rounded-full blur-3xl" />
          <div className="absolute -bottom-24 -right-16 w-80 h-80 bg-amber-500/20 rounded-full blur-3xl" />
          <Compass size={220} className="absolute -bottom-14 -right-14 text-white/[0.04] rotate-12" />
          <Anchor size={120} className="absolute top-[6%] -left-8 text-white/[0.05] -rotate-12" />
          <Cloud size={70} className="absolute top-[12%] left-[10%] text-white/20" />
          <Cloud size={46} className="absolute top-[22%] right-[14%] text-white/15" />
          <Cloud size={54} className="absolute bottom-[18%] left-[16%] text-white/10" />
          {/* Dashed flight path with a plane at the tip */}
          <svg className="absolute top-[8%] left-[8%] w-[84%] h-40 opacity-60" viewBox="0 0 600 140" fill="none">
            <path d="M10 120 C 160 20, 380 20, 560 60" stroke="#C9973B" strokeWidth="2" strokeDasharray="6 8" strokeLinecap="round" />
            <circle cx="10" cy="120" r="4" fill="#C9973B" />
          </svg>
          <Plane size={26} className="absolute top-[15%] right-[10%] text-white/70 rotate-45 animate-pulse" style={{ animation: "pdmFlyAcross 1.1s ease-out both, pulse 2.5s ease-in-out infinite 1.1s" }} />
        </div>

        <div className="relative w-full max-w-sm" style={{ animation: "pdmFadeIn .4s ease-out both" }}>
          {/* Eyebrow route strip */}
          <div className="flex items-center justify-center gap-2 mb-4 text-amber-300/90 text-[11px] font-semibold tracking-[0.2em] uppercase">
            <Sparkles size={12} />
            Perla Di Mare Travel
            <Sparkles size={12} />
          </div>

          {/* Boarding-pass card */}
          <div className="relative bg-white rounded-3xl shadow-2xl shadow-black/30 overflow-hidden" style={{ animation: "pdmPopIn .45s cubic-bezier(0.16,1,0.3,1) .1s both" }}>
            {/* Branded stub */}
            <div className="relative bg-gradient-to-r from-teal-800 to-teal-900 px-6 pt-9 pb-8 text-center overflow-hidden">
              <Plane size={90} className="absolute -bottom-4 -left-6 text-white/10 rotate-12" />
              <MapPin size={54} className="absolute top-3 right-3 text-white/10" />
              <div className="relative w-full mx-auto rounded-2xl bg-white shadow-lg flex items-center justify-center mb-3 p-4">
                <img src={LOGO_DATA_URL} alt="Perla Di Mare" className="w-full h-auto object-contain" />
              </div>
              <h1 className="relative text-white font-semibold text-lg tracking-tight" style={{ fontFamily: "'Fraunces', serif" }}>Travel Agency Manager</h1>
              <p className="relative text-teal-200/70 text-[11px] mt-0.5">By Fady Habib</p>
              <p className="relative text-teal-50/90 text-xs mt-1">Sign in to manage tickets, sales &amp; bookings</p>

              {/* Route code, like a boarding pass stub */}
              <div className="relative mt-4 flex items-center justify-center gap-3 text-white/80">
                <span className="text-sm font-bold tracking-widest">CAI</span>
                <span className="flex-1 max-w-[70px] h-px bg-white/30 relative">
                  <Plane size={12} className="absolute -top-1.5 left-1/2 -translate-x-1/2 rotate-90 text-amber-300" />
                </span>
                <span className="text-sm font-bold tracking-widest">ANY</span>
              </div>
            </div>

            {/* Perforated tear line between stub and form */}
            <div className="relative h-0">
              <div className="absolute -left-2.5 -top-2.5 w-5 h-5 rounded-full bg-teal-900" />
              <div className="absolute -right-2.5 -top-2.5 w-5 h-5 rounded-full bg-teal-900" />
              <div className="absolute left-4 right-4 top-0 border-t-2 border-dashed border-stone-200" />
            </div>

            {/* Form section */}
            <div className="relative bg-white px-6 pt-7 pb-6">
              {loginError && <div className="bg-red-50 text-red-700 text-sm rounded-2xl px-3 py-2 mb-3">{loginError}</div>}
              <div className="space-y-3">
                <div>
                  <label className="text-xs text-stone-500 block mb-1">Username</label>
                  <div className="relative">
                    <User size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
                    <input className="w-full border border-stone-300 rounded-2xl pl-9 pr-3 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-800 focus:border-teal-800"
                      value={loginUsername} onChange={(e) => setLoginUsername(e.target.value)}
                      onKeyDown={(e) => e.key === "Enter" && handleLogin()} placeholder="Username" autoFocus />
                  </div>
                </div>
                <div>
                  <label className="text-xs text-stone-500 block mb-1">Password</label>
                  <div className="relative">
                    <Lock size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
                    <input type={showPassword ? "text" : "password"}
                      className="w-full border border-stone-300 rounded-2xl pl-9 pr-9 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-800 focus:border-teal-800"
                      value={loginPassword} onChange={(e) => setLoginPassword(e.target.value)}
                      onKeyDown={(e) => e.key === "Enter" && handleLogin()} placeholder="Password" />
                    <button type="button" onClick={() => setShowPassword(!showPassword)}
                      className="absolute right-2 top-1/2 -translate-y-1/2 text-stone-400 hover:text-stone-600">
                      {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                    </button>
                  </div>
                </div>
              </div>
              <button onClick={handleLogin}
                className="group w-full mt-5 bg-gradient-to-r from-teal-800 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-2xl px-4 py-2.5 flex items-center justify-center gap-2 shadow-lg shadow-teal-800/30 transition-all">
                Sign in
                <Plane size={15} className="rotate-45 transition-transform group-hover:translate-x-1 group-hover:-translate-y-1" />
              </button>
              <p className="text-xs text-stone-400 mt-4 text-center flex items-center justify-center gap-1">
                <ShieldCheck size={13} /> Ask your admin if you don't have an account yet.
              </p>

              {/* Barcode flourish, echoing a real boarding pass stub */}
              <div className="flex items-end gap-[2px] justify-center mt-5 h-5 opacity-25">
                {[3,1,2,4,1,3,2,1,4,2,3,1,2,4,1,3,2,4,1,2,3,1,4,2,1,3,2,4,1,2].map((h, i) => (
                  <span key={i} className="bg-stone-900 w-[2px]" style={{ height: `${h * 4}px` }} />
                ))}
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // Builds one ticket's row(s) for the main table: a row per customer (skipping any
  // customer whose ticket number has been superseded by a reissue, unless that exact
  // row is the one currently jumped-to/highlighted), plus a refund sub-row if a refund
  // has been recorded on this specific ticket.
  const buildTicketRows = (t) => {
    const customers = getCustomers(t);
    const isMulti = customers.length > 1;
    const rows = [];
    customers.forEach((c, i) => {
      const ticketKey = `ticket:${(c.ticketNumber || "").trim().toUpperCase()}`;
      const isHighlighted = !!c.ticketNumber && highlightedRowKey === ticketKey;
      const isSuperseded = !!c.ticketNumber && supersededTicketNumbers.has(c.ticketNumber.trim().toUpperCase());
      // Superseded (reissued-away) rows are hidden from the table by default — they
      // only reappear when jumped to via the "Exchanged" badge on the newer ticket.
      if (isSuperseded && !isHighlighted) return;
      rows.push(
        <tr
          key={`${t.id}-${i}`}
          data-row-key={ticketKey}
          onClick={() => openTicketDetail(t)}
          className={`border-t leading-tight cursor-pointer ${
            isHighlighted
              ? `border-amber-300 bg-amber-100 ring-1 ring-inset ring-amber-400 hover:bg-amber-100 ${i > 0 ? "border-t-0" : ""}`
              : t.isReissued
              ? `border-sky-200 bg-sky-50 hover:bg-sky-100 ${i > 0 ? "border-t-0" : ""}`
              : `border-stone-100 ${i > 0 ? "border-t-0" : ""} ${isMulti ? "bg-amber-50 hover:bg-amber-100" : "hover:bg-teal-50/60"}`
          }`}
        >
          <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">{t.employee || "-"}</td>
          <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">{t.date ? formatDisplayDate(t.date) : "-"}</td>
          <td className="px-2.5 py-1 font-medium text-stone-800 whitespace-nowrap">
            <span className="inline-flex items-center gap-1.5">
              {c.name || "-"}
              {isMulti && i === 0 && (
                <span
                  title={`This booking has ${customers.length} customers / tickets`}
                  className="inline-flex items-center gap-0.5 text-[10px] font-semibold text-amber-700 bg-amber-100 border border-amber-300 rounded-full px-1.5 py-0.5"
                >
                  <Users size={10} /> {customers.length}
                </span>
              )}
            </span>
          </td>
          <td className="px-2.5 py-1 text-stone-600 font-mono whitespace-nowrap">
            <span className="inline-flex items-center gap-1.5">
              {c.ticketNumber || "-"}
              {t.isReissued && (
                <span
                  onClick={(e) => {
                    e.stopPropagation();
                    jumpToRow(`ticket:${(t.oldTicketNumber || "").trim().toUpperCase()}`);
                  }}
                  title={`Exchanged from ${t.oldTicketNumber || "an older ticket"} — click to view the original ticket`}
                  className="inline-flex items-center text-[10px] font-semibold text-sky-700 bg-sky-100 border border-sky-300 rounded-full px-1.5 py-0.5 hover:bg-sky-200 cursor-pointer"
                >
                  Exchanged{t.oldTicketNumber ? ` (orig: ${t.oldTicketNumber})` : ""}
                </span>
              )}
              {refundForIndex(t, i) && (
                <span
                  onClick={(e) => {
                    e.stopPropagation();
                    jumpToRow(`refund:${(c.ticketNumber || "").trim().toUpperCase()}`);
                  }}
                  title={`Refunded — Airline: ${fmt(refundForIndex(t, i).airlineAmount)} · Customer: ${fmt(refundForIndex(t, i).customerAmount)} — click to view the refund`}
                  className="inline-flex items-center text-[10px] font-semibold text-red-700 bg-red-100 border border-red-300 rounded-full px-1.5 py-0.5 hover:bg-red-200 cursor-pointer"
                >
                  Refunded
                </span>
              )}
            </span>
          </td>
          <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap" title={getAirlineNameByIata(t.airline) || t.airline || ""}>
            {t.airline ? (getAirlineIata(t.airline) || t.airline) : "-"}
          </td>
          <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">{routeLabel(t)}</td>
          <td className="px-2.5 py-1 text-stone-600 text-right whitespace-nowrap">{fmt(t.netPrice)}</td>
          <td className="px-2.5 py-1 text-stone-600 text-right whitespace-nowrap">{fmt(t.soldPrice)}</td>
          <td className="px-2.5 py-1 font-semibold text-emerald-700 text-right whitespace-nowrap">{fmt(profit(t.netPrice, t.soldPrice))}</td>
          <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">
            {t.company && t.company.trim() ? (
              t.company
            ) : (
              <span className="text-stone-400 italic">Individual</span>
            )}
          </td>
          <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">{t.supplier || "-"}</td>
        </tr>
      );
    });
    getRefunds(t)
      .filter((r) => r && (r.airlineAmount !== "" || r.customerAmount !== ""))
      .forEach((refund, ri) => {
      const refundedCustomer = customers[refund.customerIndex || 0];
      const refundTicketNumber = (refundedCustomer && refundedCustomer.ticketNumber) || (customers[0] && customers[0].ticketNumber) || "-";
      const refundKey = `refund:${(refundTicketNumber || "").trim().toUpperCase()}`;
      const isHighlighted = highlightedRowKey === refundKey;
      rows.push(
        <tr
          key={`${t.id}-refund-${ri}`}
          data-row-key={refundKey}
          onClick={() => openTicketDetail(t)}
          className={`border-t border-dashed leading-tight cursor-pointer ${
            isHighlighted
              ? "border-amber-300 bg-amber-100 ring-1 ring-inset ring-amber-400 hover:bg-amber-100"
              : "border-red-200 bg-red-50/60 hover:bg-red-100/60"
          }`}
        >
          <td className="px-2.5 py-1 text-red-700 whitespace-nowrap">{t.employee || "-"}</td>
          <td className="px-2.5 py-1 text-red-700 whitespace-nowrap">{refund.date ? formatDisplayDate(refund.date) : "-"}</td>
          <td className="px-2.5 py-1 font-medium text-red-800 whitespace-nowrap">{(refundedCustomer && refundedCustomer.name) || "-"}</td>
          <td className="px-2.5 py-1 text-red-700 font-mono whitespace-nowrap">
            <span className="inline-flex items-center gap-1.5">
              {refundTicketNumber}
              <span
                onClick={(e) => {
                  e.stopPropagation();
                  jumpToRow(`ticket:${(refundTicketNumber || "").trim().toUpperCase()}`);
                }}
                title="Click to view the original ticket"
                className="inline-flex items-center text-[10px] font-semibold text-red-700 bg-red-100 border border-red-300 rounded-full px-1.5 py-0.5 hover:bg-red-200 cursor-pointer"
              >
                ↳ Refund
              </span>
            </span>
          </td>
          <td className="px-2.5 py-1 text-red-700 whitespace-nowrap" title={getAirlineNameByIata(t.airline) || t.airline || ""}>
            {t.airline ? (getAirlineIata(t.airline) || t.airline) : "-"}
          </td>
          <td className="px-2.5 py-1 text-red-700 whitespace-nowrap">{routeLabel(t)}</td>
          <td className="px-2.5 py-1 text-red-700 text-right whitespace-nowrap">{fmt(refund.airlineAmount)}</td>
          <td className="px-2.5 py-1 text-red-700 text-right whitespace-nowrap">{fmt(refund.customerAmount)}</td>
          <td className="px-2.5 py-1 font-semibold text-red-800 text-right whitespace-nowrap">
            {fmt((parseFloat(refund.airlineAmount) || 0) - (parseFloat(refund.customerAmount) || 0))}
          </td>
          <td className="px-2.5 py-1 text-red-700 whitespace-nowrap">
            {t.company && t.company.trim() ? t.company : <span className="text-red-400 italic">Individual</span>}
          </td>
          <td className="px-2.5 py-1 text-red-700 whitespace-nowrap">{t.supplier || "-"}</td>
        </tr>
      );
    });
    return rows;
  };

  // ---------- Render: main app ----------
  return (
    <div
      dir="ltr"
      className="w-full min-h-screen bg-gradient-to-b from-stone-50 via-white to-teal-50/50 text-stone-800 anim-fade-in"
      style={{ fontFamily: "'Inter', sans-serif" }}
    >
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:ital,opsz,wght@0,9..144,500;0,9..144,600;1,9..144,500&family=Inter:wght@400;500;600;700&display=swap');
        .price-input::-webkit-outer-spin-button,
        .price-input::-webkit-inner-spin-button {
          -webkit-appearance: none;
          margin: 0;
        }
        .price-input[type=number] {
          -moz-appearance: textfield;
        }

        /* ---------- Global app animations ---------- */
        @keyframes pdmFadeIn { from { opacity: 0; } to { opacity: 1; } }
        @keyframes pdmPopIn {
          from { opacity: 0; transform: translateY(10px) scale(0.96); }
          to { opacity: 1; transform: translateY(0) scale(1); }
        }
        @keyframes pdmSlideUp {
          from { opacity: 0; transform: translateY(16px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes pdmSlideRight {
          from { opacity: 0; transform: translateX(24px); }
          to { opacity: 1; transform: translateX(0); }
        }
        @keyframes pdmSpin { to { transform: rotate(360deg); } }

        .anim-fade-in { animation: pdmFadeIn 0.35s ease-out both; }
        .anim-slide-up { animation: pdmSlideUp 0.4s cubic-bezier(0.16, 1, 0.3, 1) both; }
        .anim-slide-right { animation: pdmSlideRight 0.3s cubic-bezier(0.16, 1, 0.3, 1) both; }
        .anim-spin-slow { animation: pdmSpin 1.4s linear infinite; }

        @media (prefers-reduced-motion: no-preference) {
          /* The three "Totals" summary cards (Tickets / Total sales / Total profit,
             and their equivalents on the other tabs) slide up into place whenever
             the tab is switched or the filters change, and lift slightly on hover
             so the page feels a little more alive. */
          div[class*="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0"] {
            animation: pdmSlideUp 0.35s cubic-bezier(0.16, 1, 0.3, 1) both;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
          }
          div[class*="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0"]:hover {
            transform: translateY(-3px);
            box-shadow: 0 10px 20px -8px rgba(15, 118, 110, 0.25);
          }
          /* Stagger the three cards slightly so they don't all pop in at once */
          div[class*="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0"]:nth-child(2) {
            animation-delay: 0.05s;
          }
          div[class*="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0"]:nth-child(3) {
            animation-delay: 0.1s;
          }

          /* Section tabs (Flights/Hotels/Visa/...) lift gently on hover */
          button[class*="shrink-0 flex flex-col items-center gap-1.5 px-4 md:px-6 py-2.5 md:py-3 rounded-2xl border"] {
            transition: transform 0.15s ease, box-shadow 0.15s ease, background-color 0.15s ease, border-color 0.15s ease, color 0.15s ease;
          }
          button[class*="shrink-0 flex flex-col items-center gap-1.5 px-4 md:px-6 py-2.5 md:py-3 rounded-2xl border"]:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 14px -6px rgba(15, 118, 110, 0.3);
          }
        }

        @media (prefers-reduced-motion: no-preference) {
          /* Any full-screen overlay (modal backdrops and full-page panels) fades in */
          div[class*="fixed inset-0"] {
            animation: pdmFadeIn 0.18s ease-out both;
          }
          /* The modal card itself (always the overlay's first child) pops/slides in */
          div[class*="fixed inset-0"][class*="bg-black/40"] > div:first-child,
          div[class*="fixed inset-0"][class*="bg-stone-900/40"] > div:first-child {
            animation: pdmPopIn 0.24s cubic-bezier(0.16, 1, 0.3, 1) both;
          }

          /* Buttons and clickable controls get soft press/hover feedback */
          button:not(:disabled),
          [role="button"] {
            transition: transform 0.15s ease, box-shadow 0.15s ease,
              background-color 0.15s ease, border-color 0.15s ease,
              color 0.15s ease, opacity 0.15s ease !important;
          }
          button:not(:disabled):active,
          [role="button"]:active {
            transform: scale(0.96);
          }

          /* Inputs animate their focus ring smoothly */
          input, select, textarea {
            transition: box-shadow 0.15s ease, border-color 0.15s ease, background-color 0.15s ease;
          }
        }

        @media (prefers-reduced-motion: reduce) {
          *, *::before, *::after {
            animation-duration: 0.001ms !important;
            animation-iteration-count: 1 !important;
            transition-duration: 0.001ms !important;
          }
        }
      `}</style>
      <div className="xl:flex xl:justify-center xl:gap-4">
      <div className="relative max-w-5xl mx-auto xl:mx-0 xl:w-[64rem] xl:shrink-0 p-4 md:p-6">
        {/* Boarding-pass style banner */}
        <div className="relative rounded-2xl bg-gradient-to-r from-teal-800 via-teal-800 to-teal-900 shadow-lg shadow-teal-900/20 overflow-hidden mb-0">
          <Plane size={140} className="pointer-events-none absolute -bottom-8 -right-6 text-white/[0.06] rotate-45" />
          <Compass size={90} className="pointer-events-none absolute -top-6 left-[38%] text-white/[0.05]" />
          <Luggage size={70} className="pointer-events-none absolute -bottom-4 left-[18%] text-white/[0.05] hidden md:block" />
          <header className="relative flex items-center justify-between flex-wrap gap-3 px-4 py-4 md:px-6">
            <div className="flex items-center gap-3">
              <div className="bg-white rounded-2xl p-2.5 shadow-sm shrink-0 hidden sm:block">
                <img src={LOGO_DATA_URL} alt="Perla Di Mare" className="w-[120px] h-auto md:w-[150px] object-contain" />
              </div>
              <div>
                <h1 className="text-lg md:text-2xl font-semibold text-white" style={{ fontFamily: "'Fraunces', serif" }}>
                  Travel Agency Manager <span className="text-teal-200/60 font-medium text-xs md:text-base" style={{ fontFamily: "'Inter', sans-serif" }}>By Fady Habib</span>
                </h1>
                <p className="text-teal-100/80 text-sm flex items-center gap-1.5 flex-wrap mt-0.5">
                  Signed in as {currentUser.name}
                  {currentUser.isAdmin && (
                    <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-teal-900 bg-amber-300 border border-amber-400/50 rounded-full px-2 py-0.5">
                      <ShieldCheck size={11} /> Main account
                    </span>
                  )}
                  {hasAdminAccess && !currentUser.isAdmin && currentEmployeeRecord && (
                    <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-white bg-white/10 border border-white/20 rounded-full px-2 py-0.5">
                      {roleLabel(currentEmployeeRecord.role)}
                    </span>
                  )}
                  {hasAdminAccess && isAccountingUser && (
                    <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-amber-100 bg-amber-500/20 border border-amber-300/30 rounded-full px-2 py-0.5">
                      Accounting — view only
                    </span>
                  )}
                  {hasAdminAccess && (
                    <button
                      type="button"
                      onClick={() => setShowOnlineList(!showOnlineList)}
                      className="inline-flex items-center gap-1 text-[10px] font-semibold text-emerald-50 bg-emerald-500/20 border border-emerald-300/30 rounded-full px-2 py-0.5 hover:bg-emerald-500/30"
                    >
                      <Wifi size={11} />
                      {visibleOnlineUsernames.length} online now
                    </button>
                  )}
                </p>
              </div>
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              {activeSection === "accounts" && canAccessAccounts && (
                <button
                  type="button"
                  onClick={() => setAccountsLang((l) => (l === "ar" ? "en" : "ar"))}
                  title={accountsLang === "ar" ? "Switch to English" : "التحويل للعربية"}
                  className="border border-white/20 bg-white/10 hover:bg-white/20 text-white text-sm rounded-2xl p-2 flex items-center justify-center gap-1 transition-colors"
                >
                  <Globe size={15} />
                  <span className="text-xs font-semibold">{accountsLang === "ar" ? "EN" : "AR"}</span>
                </button>
              )}
              {hasAdminAccess && (
                <button onClick={handleBackup} title="Backup"
                  className="border border-white/20 bg-white/10 hover:bg-white/20 text-white text-sm rounded-2xl p-2 flex items-center justify-center transition-colors">
                  <Download size={15} />
                </button>
              )}
              {hasAdminAccess && (
                <button onClick={triggerRestore} title="Restore"
                  className="border border-white/20 bg-white/10 hover:bg-white/20 text-white text-sm rounded-2xl p-2 flex items-center justify-center transition-colors">
                  <Upload size={15} />
                </button>
              )}
              {hasAdminAccess && (
                <input
                  type="file"
                  accept="application/json"
                  ref={fileInputRef}
                  onChange={handleRestoreFile}
                  className="hidden"
                />
              )}
              {hasAdminAccess && (
                <button onClick={() => setShowManage(!showManage)} title="Manage employees"
                  className="border border-white/20 bg-white/10 hover:bg-white/20 text-white text-sm rounded-2xl p-2 flex items-center justify-center transition-colors">
                  <Users size={15} />
                </button>
              )}
              {currentUser.isAdmin && (
                <button onClick={() => setShowLicensePanel(!showLicensePanel)} title={isLicensed ? "License" : "Activate license"}
                  className={`border text-sm rounded-2xl p-2 flex items-center justify-center transition-colors ${
                    isLicensed
                      ? "border-white/20 bg-white/10 hover:bg-white/20 text-white"
                      : "border-amber-300/50 bg-amber-500/20 hover:bg-amber-500/30 text-amber-100"
                  }`}>
                  <Lock size={15} />
                </button>
              )}
              {canManageCompanies && (
                <button onClick={() => setShowManageCompanies(!showManageCompanies)} title="Manage companies"
                  className="border border-white/20 bg-white/10 hover:bg-white/20 text-white text-sm rounded-2xl p-2 flex items-center justify-center transition-colors">
                  <Factory size={15} />
                </button>
              )}
              <button
                onClick={() => {
                  setShowRequestsPanel(!showRequestsPanel);
                  setRequestSendError("");
                }}
                title="Requests"
                className="relative border border-white/20 bg-white/10 hover:bg-white/20 text-white text-sm rounded-2xl p-2 flex items-center justify-center transition-colors">
                <Bell size={15} />
                {myPendingRequestsCount > 0 && (
                  <span className="absolute -top-1.5 -right-1.5 min-w-[18px] h-[18px] px-1 rounded-full bg-red-500 text-white text-[10px] font-bold flex items-center justify-center">
                    {myPendingRequestsCount}
                  </span>
                )}
              </button>
              <button
                onClick={() => {
                  setShowChangePassword(!showChangePassword);
                  setPasswordError("");
                  setPasswordSuccess("");
                  setCurrentPasswordInput("");
                  setNewPasswordInput("");
                  setConfirmPasswordInput("");
                }}
                title="Change password"
                className="border border-white/20 bg-white/10 hover:bg-white/20 text-white text-sm rounded-2xl p-2 flex items-center justify-center transition-colors">
                <Lock size={15} />
              </button>
              <button onClick={handleLogout} title="Sign out"
                className="border border-white/20 bg-white/10 hover:bg-white/20 text-white text-sm rounded-2xl p-2 flex items-center justify-center transition-colors">
                <LogOut size={15} />
              </button>
              {onChangeServer && (
                <button
                  onClick={() => {
                    requestConfirm(
                      `Disconnect from the current server${currentServerUrl ? ` (${currentServerUrl})` : ""} and connect to a different one?`,
                      () => {
                        setConfirmDialog(null);
                        onChangeServer();
                      }
                    );
                  }}
                  title="Server"
                  className="border border-white/20 bg-white/10 hover:bg-white/20 text-teal-100 text-sm rounded-2xl p-2 flex items-center justify-center transition-colors"
                >
                  <Wifi size={15} />
                </button>
              )}
            </div>
          </header>
        </div>
        {/* Perforated tear line, like separating a boarding-pass stub from the rest */}
        <div className="relative h-6 mb-4">
          <div className="absolute -left-2.5 top-0 w-5 h-5 rounded-full bg-stone-50" />
          <div className="absolute -right-2.5 top-0 w-5 h-5 rounded-full bg-stone-50" />
          <div className="absolute left-4 right-4 top-2.5 border-t-2 border-dashed border-teal-800/20" />
        </div>

        {showLicensePanel && currentUser.isAdmin && (
          <div
            className="fixed inset-0 z-50 bg-black/40 flex items-start md:items-center justify-center p-4 overflow-y-auto"
            onClick={(e) => {
              if (e.target === e.currentTarget) {
                setShowLicensePanel(false);
                setLicenseError("");
                setLicenseInput("");
              }
            }}
          >
            <div className="bg-white rounded-2xl border border-stone-200 p-4 md:p-5 w-full max-w-sm my-8 md:my-0 max-h-[90vh] overflow-y-auto">
              <div className="flex items-center justify-between mb-1">
                <h2 className="font-semibold text-stone-900 flex items-center gap-2">
                  <Lock size={16} className="text-teal-800" /> App license
                </h2>
                <button
                  onClick={() => { setShowLicensePanel(false); setLicenseError(""); setLicenseInput(""); }}
                  className="text-stone-400 hover:text-stone-600 p-1 -m-1 rounded-lg hover:bg-stone-100"
                >
                  <X size={18} />
                </button>
              </div>
              {isLicensed ? (
                <div className="bg-emerald-50 text-emerald-700 text-sm rounded-xl px-3 py-2 mb-4 mt-3">
                  Active{licenseRecord && licenseRecord.expiresAt ? ` — valid until ${licenseRecord.expiresAt}` : " — permanent license"}
                </div>
              ) : (
                <p className="text-xs text-stone-400 mb-4 mt-3">
                  {licenseRecord ? "The current activation code has expired." : "The app is not activated yet."} Enter a valid activation code below — the app stays locked for every employee until this is done.
                </p>
              )}
              {isLicensed && (
                <p className="text-xs text-stone-400 mb-4">
                  Entering a new code below will replace the current one — useful for renewing or switching licenses.
                </p>
              )}
              {licenseError && <div className="bg-red-50 text-red-700 text-sm rounded-xl px-3 py-2 mb-3">{licenseError}</div>}
              <div>
                <label className="text-xs text-stone-500 block mb-1">Activation code</label>
                <input
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 tracking-widest uppercase"
                  value={licenseInput}
                  onChange={(e) => setLicenseInput(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && handleActivateLicense()}
                  placeholder="XXXX-XXXX-XXXX"
                  autoFocus
                />
              </div>
              <button
                onClick={handleActivateLicense}
                disabled={licenseSaving}
                className="w-full mt-4 bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 shadow-sm shadow-teal-800/30 ring-1 ring-inset ring-white/10 transition-colors disabled:opacity-60">
                {licenseSaving ? "Saving..." : "Activate"}
              </button>
            </div>
          </div>
        )}

        {(showManage || showManageCompanies) && (
          <div
            className="fixed inset-0 z-50 bg-black/40 flex items-start md:items-center justify-center p-4 overflow-y-auto"
            onClick={(e) => {
              if (e.target === e.currentTarget) {
                setShowManage(false);
                setShowManageCompanies(false);
              }
            }}
          >
          <div className="bg-stone-50 rounded-2xl w-full max-w-3xl my-8 md:my-0 max-h-[90vh] overflow-y-auto p-1" onClick={(e) => e.stopPropagation()}>
        {showManage && hasAdminAccess && (
          <div className="bg-stone-50">
            <button onClick={() => setShowManage(false)}
              className="mb-4 border border-stone-300 text-stone-600 text-sm rounded-xl px-3 py-2 flex items-center gap-1.5 hover:bg-stone-100">
              <X size={15} /> Close
            </button>
          <div className="bg-white rounded-2xl border border-stone-200 p-4 md:p-5 mb-6">
            <h2 className="font-semibold text-stone-900 mb-1">Employee accounts</h2>
            <p className="text-xs text-stone-400 mb-4">
              As the main account, you can view and change every employee's password, edit their name or username, add or remove accounts, and grant or remove main-account access. A grade (Manager, Supervisor, Employee, Accountant, Owner — plus a per-department version of Manager/Supervisor/Employee) is chosen once when the employee is first added, from the three grade dropdowns on the Add employee card below, and fills in a starting set of permissions. Every permission — view all services, add, edit, delete, accounting/notes-only mode, manage companies, Owner access, and which sections (Flights, Hotels, Visa, Transportation, Files) they can access — stays an individual on/off switch you can set by hand for each employee afterward: click their name to open it. An Owner gets everything a main account has (Manage employees, Backup/Restore, every ticket permission) except the License panel, which stays reserved for main accounts. This is a basic access gate, not a secure authentication system — anyone with technical access to the app's stored data can read these passwords. Avoid reusing important passwords here.
            </p>
            {manageError && <div className="bg-red-50 text-red-700 text-sm rounded-xl px-3 py-2 mb-3">{manageError}</div>}
            <p className="text-xs text-stone-500 mb-3 flex items-center gap-1.5">
              <Wifi size={13} className="text-emerald-600" />
              {visibleOnlineUsernames.length} of {(employees || []).filter((e) => currentUser.isAdmin || !e.isAdmin).length} employees connected right now
            </p>
            <div className="border border-stone-200 rounded-xl overflow-hidden mb-4">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-stone-50 text-stone-500 text-xs">
                    <th className="text-left px-3 py-2 font-medium">Status</th>
                    <th className="text-left px-3 py-2 font-medium">Name</th>
                    <th className="text-left px-3 py-2 font-medium">Username</th>
                    <th className="text-left px-3 py-2 font-medium">Password</th>
                    <th className="text-left px-3 py-2 font-medium">Grade</th>
                    <th className="text-left px-3 py-2 font-medium"></th>
                  </tr>
                </thead>
                <tbody>
                  {(employees || [])
                    .filter((e) => currentUser.isAdmin || !e.isAdmin)
                    .map((e) => {
                    const isEditing = editingUsername === e.username;
                    if (isEditing) {
                      return (
                        <tr key={e.username} className="border-t border-stone-100 bg-stone-50">
                          <td className="px-3 py-2">
                            <span className={`inline-flex items-center gap-1 text-[10px] font-semibold rounded-full px-2 py-0.5 ${isOnline(e.username) ? "text-emerald-700 bg-emerald-50 border border-emerald-200" : "text-stone-400 bg-stone-100 border border-stone-200"}`}>
                              <span className={`w-1.5 h-1.5 rounded-full ${isOnline(e.username) ? "bg-emerald-500" : "bg-stone-300"}`} />
                              {isOnline(e.username) ? "Online" : "Offline"}
                            </span>
                          </td>
                          <td className="px-3 py-2">
                            <input
                              className="w-full border border-stone-300 rounded-xl px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                              value={editDraft.name}
                              onChange={(ev) => setEditDraft({ ...editDraft, name: ev.target.value })}
                            />
                          </td>
                          <td className="px-3 py-2">
                            <input
                              className="w-full border border-stone-300 rounded-xl px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                              value={editDraft.username}
                              onChange={(ev) => setEditDraft({ ...editDraft, username: ev.target.value })}
                            />
                          </td>
                          <td className="px-3 py-2">
                            <div className="relative">
                              <input
                                type={editShowPassword ? "text" : "password"}
                                className="w-full border border-stone-300 rounded-xl pl-2 pr-8 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                                value={editDraft.password}
                                onChange={(ev) => setEditDraft({ ...editDraft, password: ev.target.value })}
                              />
                              <button
                                type="button"
                                onClick={() => setEditShowPassword(!editShowPassword)}
                                className="absolute right-2 top-1/2 -translate-y-1/2 text-stone-400"
                              >
                                {editShowPassword ? <EyeOff size={14} /> : <Eye size={14} />}
                              </button>
                            </div>
                          </td>
                          <td className="px-3 py-2 text-stone-500">
                            {e.isAdmin ? (
                              <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-teal-800 bg-teal-50 border border-teal-200 rounded-full px-2 py-0.5">
                                <ShieldCheck size={11} /> Main
                              </span>
                            ) : (
                              <span className="text-xs font-medium text-stone-600">{roleLabel(e.role)}</span>
                            )}
                          </td>
                          <td className="px-3 py-2">
                            <div className="flex gap-1 justify-end">
                              <button onClick={saveEditEmployee} className="text-emerald-600 hover:text-emerald-800 p-1">
                                <Check size={15} />
                              </button>
                              <button onClick={cancelEditEmployee} className="text-stone-400 hover:text-red-600 p-1">
                                <X size={15} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    }
                    return (
                      <tr key={e.username} className="border-t border-stone-100">
                        <td className="px-3 py-2">
                          <span className={`inline-flex items-center gap-1 text-[10px] font-semibold rounded-full px-2 py-0.5 ${isOnline(e.username) ? "text-emerald-700 bg-emerald-50 border border-emerald-200" : "text-stone-400 bg-stone-100 border border-stone-200"}`}>
                            <span className={`w-1.5 h-1.5 rounded-full ${isOnline(e.username) ? "bg-emerald-500" : "bg-stone-300"}`} />
                            {isOnline(e.username) ? "Online" : "Offline"}
                          </span>
                        </td>
                        <td className="px-3 py-2">
                          {e.isAdmin ? (
                            e.name
                          ) : (
                            <button
                              type="button"
                              onClick={() => setOpenPermissionsFor(e.username)}
                              title="View or change this employee's permissions"
                              className="text-teal-800 hover:text-teal-900 hover:underline font-medium text-left"
                            >
                              {e.name}
                            </button>
                          )}
                        </td>
                        <td className="px-3 py-2 text-stone-500">{e.username}</td>
                        <td className="px-3 py-2 text-stone-500">
                          <span className="font-mono" title="Passwords are stored securely and can't be viewed — click the employee's name to set a new one">••••••••</span>
                        </td>
                        <td className="px-3 py-2 text-stone-500">
                          {e.isAdmin ? (
                            <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-teal-800 bg-teal-50 border border-teal-200 rounded-full px-2 py-0.5">
                              <ShieldCheck size={11} /> Main
                            </span>
                          ) : (
                            <span className="text-xs font-medium text-stone-600">{roleLabel(e.role)}</span>
                          )}
                        </td>
                        <td className="px-3 py-2 text-right">
                          <div className="flex gap-1 justify-end">
                            {/* Promoting/demoting main-account access stays reserved for true
                                main accounts — an Owner never gets to touch this, so an Owner
                                can never grant themselves (or anyone else) admin access and
                                route around the License restriction. Editing the grade, and
                                editing/deleting the employee entirely, now live inside the
                                Permissions modal (opened via the name link) instead of here. */}
                            {e.isAdmin ? (
                              currentUser.isAdmin && (
                                <button
                                  onClick={() => handleDemoteAdmin(e.username)}
                                  title="Remove main-account access"
                                  className="text-stone-400 hover:text-amber-600 text-[11px] font-semibold border border-stone-200 rounded-lg px-1.5 py-1"
                                >
                                  Remove main
                                </button>
                              )
                            ) : (
                              currentUser.isAdmin && (
                                <button
                                  onClick={() => handlePromoteToAdmin(e.username)}
                                  title="Make this a main account"
                                  className="text-stone-400 hover:text-teal-800 text-[11px] font-semibold border border-stone-200 rounded-lg px-1.5 py-1 flex items-center gap-1"
                                >
                                  <ShieldCheck size={12} /> Make main
                                </button>
                              )
                            )}
                            {e.isAdmin && currentUser.isAdmin && (
                              <button onClick={() => startEditEmployee(e)} className="text-stone-400 hover:text-teal-800 p-1">
                                <Pencil size={15} />
                              </button>
                            )}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            {/* Whole "add a new employee" block lives in one card: name/username/
                password, the grade picker, and the Add button. */}
            <div className="border border-stone-200 rounded-2xl p-4 bg-stone-50">
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                <input className="border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white"
                  placeholder="Full name" value={newEmployee.name} autoComplete="off"
                  onChange={(e) => setNewEmployee({ ...newEmployee, name: e.target.value })} />
                <input className="border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white"
                  placeholder="Username" value={newEmployee.username} autoComplete="off" name="new-employee-username"
                  onChange={(e) => setNewEmployee({ ...newEmployee, username: e.target.value })} />
                <input type="password" className="border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white"
                  placeholder="Password" value={newEmployee.password} autoComplete="new-password" name="new-employee-password"
                  onChange={(e) => setNewEmployee({ ...newEmployee, password: e.target.value })} />
              </div>

              {/* Grade picker: one dropdown per tier (Manager / Supervisor / Employee),
                  each holding that tier's general grade plus its per-department variants.
                  Owner, GM, and Accountant stand alone next to the three dropdowns since
                  none of them has per-department variants. Picking any option sets that grade's
                  starting permissions on newEmployee and immediately closes its dropdown,
                  so the chosen grade shows outside/above the list. Name/username/password
                  above are never touched by any of this — they only get cleared once the
                  employee is actually added (see handleAddEmployee), so they stay exactly
                  as typed while you pick a grade. Everything can still be fine-tuned
                  afterward from the Permissions screen reached by clicking the employee's
                  name once they've been added. */}
              <div className="mt-3">
                <label className="text-xs text-stone-500 block mb-1.5">Grade</label>
                <div className="flex flex-wrap items-start gap-1.5 pb-1">
                  {GRADE_TIER_GROUPS.map((group) => {
                    const selectedInGroup = group.roles.find((r) => r.value === newEmployee.role);
                    const isOpen = newEmployeeGradeOpen === group.key;
                    return (
                      <div key={group.key} className="relative w-28 shrink-0">
                        <button
                          type="button"
                          onClick={() => setNewEmployeeGradeOpen(isOpen ? null : group.key)}
                          className={`w-full flex items-center justify-between border rounded-xl px-2 py-2 text-xs bg-white hover:bg-stone-50 focus:outline-none focus:ring-2 focus:ring-teal-700 ${
                            selectedInGroup ? "border-teal-700" : "border-stone-300"
                          }`}
                        >
                          <span className={`font-medium truncate ${selectedInGroup ? "text-teal-800" : "text-stone-700"}`}>
                            {selectedInGroup ? selectedInGroup.label : group.title}
                          </span>
                          <ChevronDown size={13} className={`text-stone-400 shrink-0 transition-transform ${isOpen ? "rotate-180" : ""}`} />
                        </button>

                        {isOpen && (
                          <div className="absolute z-10 mt-1 w-56 border border-stone-200 rounded-xl bg-white shadow-lg max-h-64 overflow-y-auto">
                            {group.roles.map((r) => {
                              const selected = newEmployee.role === r.value;
                              return (
                                <button
                                  key={r.value}
                                  type="button"
                                  onClick={() => {
                                    setNewEmployee({ ...newEmployee, role: r.value, ...ROLE_PRESETS[r.value] });
                                    setNewEmployeeGradeOpen(null);
                                  }}
                                  className="w-full flex items-center gap-2 px-3 py-2 text-sm text-left hover:bg-stone-50 border-b border-stone-100 last:border-b-0"
                                >
                                  <span
                                    className={`shrink-0 w-4 h-4 rounded-full border flex items-center justify-center ${
                                      selected ? "border-teal-700" : "border-stone-300"
                                    }`}
                                  >
                                    {selected && <span className="w-2 h-2 rounded-full bg-teal-700" />}
                                  </span>
                                  <span className={selected ? "text-teal-800 font-semibold" : "text-stone-600"}>{r.label}</span>
                                </button>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    );
                  })}

                  {/* Owner, GM, and Accountant have no per-department variants, so they sit
                      as standalone picks next to the three tier dropdowns rather than inside one. */}
                  {["owner", "gm", "accountant"].map((value) => {
                    const selected = newEmployee.role === value;
                    return (
                      <button
                        key={value}
                        type="button"
                        onClick={() => setNewEmployee({ ...newEmployee, role: value, ...ROLE_PRESETS[value] })}
                        className={`w-24 shrink-0 flex items-center gap-1 border rounded-xl px-2 py-2 text-xs transition-colors ${
                          selected
                            ? "bg-teal-800 text-white border-teal-800"
                            : "bg-white text-stone-600 border-stone-300 hover:bg-stone-50"
                        }`}
                      >
                        <span
                          className={`shrink-0 w-4 h-4 rounded-full border flex items-center justify-center ${
                            selected ? "border-white" : "border-stone-300"
                          }`}
                        >
                          {selected && <span className="w-2 h-2 rounded-full bg-white" />}
                        </span>
                        {roleLabel(value)}
                      </button>
                    );
                  })}
                </div>

                <p className="text-[11px] text-stone-400 mt-2">
                  You can fine-tune every permission for each section after adding them, from Manage employees.
                </p>
              </div>

              <button onClick={handleAddEmployee}
                className="mt-3 bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 shadow-sm shadow-teal-800/30 ring-1 ring-inset ring-white/10 transition-colors flex items-center gap-1.5">
                <UserPlus size={15} /> Add employee
              </button>
            </div>
          </div>
          </div>
        )}
        {showManageCompanies && canManageCompanies && (
          <div className="bg-stone-50">
            <button onClick={() => setShowManageCompanies(false)}
              className="mb-4 border border-stone-300 text-stone-600 text-sm rounded-xl px-3 py-2 flex items-center gap-1.5 hover:bg-stone-100">
              <X size={15} /> Close
            </button>
          <div className="bg-white rounded-2xl border border-stone-200 p-4 md:p-5 mb-6">
            <h2 className="font-semibold text-stone-900 mb-1 flex items-center gap-2">
              <Factory size={18} className="text-stone-500" /> Companies
            </h2>
            <p className="text-xs text-stone-400 mb-4">
              Register each company's details here so they're always available to pick from the Company field and filter, even before any ticket has been entered for them.
            </p>
            {companyError && <div className="bg-red-50 text-red-700 text-sm rounded-xl px-3 py-2 mb-3">{companyError}</div>}

            <div className="grid grid-cols-1 md:grid-cols-2 gap-3 max-w-2xl mb-3">
              <input
                className="border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                placeholder="Company name"
                value={newCompanyDraft.name}
                onChange={(e) => setNewCompanyDraft({ ...newCompanyDraft, name: e.target.value })}
              />
              <input
                className="border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                placeholder="Tax number"
                value={newCompanyDraft.taxNumber}
                onChange={(e) => setNewCompanyDraft({ ...newCompanyDraft, taxNumber: e.target.value })}
              />
              <input
                className="border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                placeholder="Commercial registration number"
                value={newCompanyDraft.commercialReg}
                onChange={(e) => setNewCompanyDraft({ ...newCompanyDraft, commercialReg: e.target.value })}
              />
              <input
                className="border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                placeholder="Phone numbers (comma separated)"
                value={newCompanyDraft.phones}
                onChange={(e) => setNewCompanyDraft({ ...newCompanyDraft, phones: e.target.value })}
              />
            </div>
            <div className="flex gap-2 mb-5">
              <button
                onClick={handleAddCompany}
                className="bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 shadow-sm shadow-teal-800/30 ring-1 ring-inset ring-white/10 transition-colors flex items-center gap-1.5 whitespace-nowrap"
              >
                {editingCompanyName ? <Check size={15} /> : <Factory size={15} />}
                {editingCompanyName ? "Save changes" : "Add company"}
              </button>
              {editingCompanyName && (
                <button
                  onClick={cancelEditCompany}
                  className="border border-stone-300 text-stone-600 text-sm rounded-xl px-4 py-2 flex items-center gap-1.5"
                >
                  <X size={15} /> Cancel
                </button>
              )}
            </div>

            <div className="flex items-center justify-between mb-3">
              <p className="text-xs text-stone-400">
                {suggestions.companies.length} compan{suggestions.companies.length === 1 ? "y" : "ies"} saved
              </p>
              <button
                onClick={() => setShowCompaniesList(!showCompaniesList)}
                className="text-teal-800 border border-teal-800 hover:bg-teal-50 text-xs font-semibold rounded-xl px-3 py-1.5 flex items-center gap-1.5"
              >
                <List size={14} /> {showCompaniesList ? "Hide companies list" : "View all companies"}
              </button>
            </div>

            {showCompaniesList && (
              suggestions.companies.length === 0 ? (
                <p className="text-sm text-stone-400">No companies saved yet</p>
              ) : (
                <div className="border border-stone-200 rounded-xl overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs">
                      <thead>
                        <tr className="bg-stone-50 text-stone-500 text-[11px] uppercase tracking-wide">
                          <th className="text-left px-3 py-2 font-semibold whitespace-nowrap">Company</th>
                          <th className="text-left px-3 py-2 font-semibold whitespace-nowrap">Tax number</th>
                          <th className="text-left px-3 py-2 font-semibold whitespace-nowrap">Commercial reg.</th>
                          <th className="text-left px-3 py-2 font-semibold whitespace-nowrap">Phone</th>
                          <th className="text-right px-3 py-2 font-semibold whitespace-nowrap"></th>
                        </tr>
                      </thead>
                      <tbody>
                        {[...suggestions.companies]
                          .sort((a, b) => companyName(a).localeCompare(companyName(b)))
                          .map((c) => {
                            const name = companyName(c);
                            const taxNumber = typeof c === "object" ? c.taxNumber : "";
                            const commercialReg = typeof c === "object" ? c.commercialReg : "";
                            const phones = typeof c === "object" && Array.isArray(c.phones) ? c.phones : [];
                            return (
                              <tr
                                key={name}
                                className={`border-t border-stone-100 ${editingCompanyName === name ? "bg-teal-50/40" : "hover:bg-stone-50"}`}
                              >
                                <td className="px-3 py-2 font-medium text-stone-800 whitespace-nowrap">{name}</td>
                                <td className="px-3 py-2 text-stone-600 whitespace-nowrap">{taxNumber || "-"}</td>
                                <td className="px-3 py-2 text-stone-600 whitespace-nowrap">{commercialReg || "-"}</td>
                                <td className="px-3 py-2 text-stone-600 whitespace-nowrap">{phones.length > 0 ? phones.join(", ") : "-"}</td>
                                <td className="px-3 py-2 text-right whitespace-nowrap">
                                  <div className="flex gap-1 justify-end">
                                    <button onClick={() => handleEditCompanyClick(c)} className="text-stone-400 hover:text-teal-800 p-0.5">
                                      <Pencil size={13} />
                                    </button>
                                    <button onClick={() => handleDeleteCompany(name)} className="text-stone-400 hover:text-red-600 p-0.5">
                                      <Trash2 size={13} />
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            );
                          })}
                      </tbody>
                    </table>
                  </div>
                </div>
              )
            )}
          </div>
          </div>
        )}
          </div>
          </div>
        )}

        {!showManage && !showManageCompanies && !showLicensePanel && (
        <>
        {isLicensed ? (
        <>
        {/* Top-level section switcher */}
        <div className="flex items-center gap-2 md:gap-3 mb-6 overflow-x-auto md:justify-center md:overflow-visible -mx-4 px-4 md:mx-0 md:px-0">
          {mySections.flights && (
          <button
            onClick={() => setActiveSection("flights")}
            className={`shrink-0 flex flex-col items-center gap-1.5 px-4 md:px-6 py-2.5 md:py-3 rounded-2xl border text-xs font-semibold transition-colors ${
              activeSection === "flights"
                ? "bg-gradient-to-b from-teal-700 to-teal-900 text-white border-teal-800 shadow-md shadow-teal-800/30 ring-1 ring-inset ring-amber-600/50"
                : "bg-white text-stone-500 border-stone-200 hover:border-amber-600 hover:text-teal-800 hover:shadow-sm"
            }`}
          >
            <Plane size={22} className="rotate-45" />
            Flights
          </button>
          )}
          {mySections.hotels && (
          <button
            onClick={() => setActiveSection("hotels")}
            className={`shrink-0 flex flex-col items-center gap-1.5 px-4 md:px-6 py-2.5 md:py-3 rounded-2xl border text-xs font-semibold transition-colors ${
              activeSection === "hotels"
                ? "bg-gradient-to-b from-teal-700 to-teal-900 text-white border-teal-800 shadow-md shadow-teal-800/30 ring-1 ring-inset ring-amber-600/50"
                : "bg-white text-stone-500 border-stone-200 hover:border-amber-600 hover:text-teal-800 hover:shadow-sm"
            }`}
          >
            <Building2 size={22} />
            Hotels
          </button>
          )}
          {mySections.visa && (
          <button
            onClick={() => setActiveSection("visa")}
            className={`shrink-0 flex flex-col items-center gap-1.5 px-4 md:px-6 py-2.5 md:py-3 rounded-2xl border text-xs font-semibold transition-colors ${
              activeSection === "visa"
                ? "bg-gradient-to-b from-teal-700 to-teal-900 text-white border-teal-800 shadow-md shadow-teal-800/30 ring-1 ring-inset ring-amber-600/50"
                : "bg-white text-stone-500 border-stone-200 hover:border-amber-600 hover:text-teal-800 hover:shadow-sm"
            }`}
          >
            <PassportIcon size={22} />
            Visa
          </button>
          )}
          {mySections.cars && (
          <button
            onClick={() => setActiveSection("cars")}
            className={`shrink-0 flex flex-col items-center gap-1.5 px-4 md:px-6 py-2.5 md:py-3 rounded-2xl border text-xs font-semibold transition-colors ${
              activeSection === "cars"
                ? "bg-gradient-to-b from-teal-700 to-teal-900 text-white border-teal-800 shadow-md shadow-teal-800/30 ring-1 ring-inset ring-amber-600/50"
                : "bg-white text-stone-500 border-stone-200 hover:border-amber-600 hover:text-teal-800 hover:shadow-sm"
            }`}
          >
            <Car size={22} />
            Transportation
          </button>
          )}
          {mySections.files && (
          <button
            onClick={() => setActiveSection("files")}
            className={`shrink-0 flex flex-col items-center gap-1.5 px-4 md:px-6 py-2.5 md:py-3 rounded-2xl border text-xs font-semibold transition-colors ${
              activeSection === "files"
                ? "bg-gradient-to-b from-teal-700 to-teal-900 text-white border-teal-800 shadow-md shadow-teal-800/30 ring-1 ring-inset ring-amber-600/50"
                : "bg-white text-stone-500 border-stone-200 hover:border-amber-600 hover:text-teal-800 hover:shadow-sm"
            }`}
          >
            <FileText size={22} />
            Files
          </button>
          )}
          {canAccessAccounts && (
          <button
            onClick={() => setActiveSection("accounts")}
            className={`shrink-0 flex flex-col items-center gap-1.5 px-4 md:px-6 py-2.5 md:py-3 rounded-2xl border text-xs font-semibold transition-colors ${
              activeSection === "accounts"
                ? "bg-gradient-to-b from-teal-700 to-teal-900 text-white border-teal-800 shadow-md shadow-teal-800/30 ring-1 ring-inset ring-amber-600/50"
                : "bg-white text-stone-500 border-stone-200 hover:border-amber-600 hover:text-teal-800 hover:shadow-sm"
            }`}
          >
            <Wallet size={22} />
            الحسابات
          </button>
          )}
        </div>

        {activeSection === "flights" && (
        <>
        {currentUser.isAdmin && (restoreError || restoreSuccess) && (
          <div className={`text-sm rounded-xl px-3 py-2 mb-4 ${restoreError ? "bg-red-50 text-red-700" : "bg-emerald-50 text-emerald-700"}`}>
            {restoreError || restoreSuccess}
          </div>
        )}



        {showChangePassword && (
          <div className="bg-white rounded-2xl border border-stone-200 p-4 md:p-5 mb-6 max-w-sm">
            <h2 className="font-semibold text-stone-900 mb-1 flex items-center gap-2">
              <Lock size={16} className="text-teal-800" /> Change your password
            </h2>
            <p className="text-xs text-stone-400 mb-4">
              Signed in as {currentUser.name} ({currentUser.username})
            </p>
            {passwordError && <div className="bg-red-50 text-red-700 text-sm rounded-xl px-3 py-2 mb-3">{passwordError}</div>}
            {passwordSuccess && <div className="bg-emerald-50 text-emerald-700 text-sm rounded-xl px-3 py-2 mb-3">{passwordSuccess}</div>}
            <div className="space-y-3">
              <div>
                <label className="text-xs text-stone-500 block mb-1">Current password</label>
                <input type="password"
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={currentPasswordInput} onChange={(e) => setCurrentPasswordInput(e.target.value)} />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">New password</label>
                <input type="password"
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={newPasswordInput} onChange={(e) => setNewPasswordInput(e.target.value)} />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Confirm new password</label>
                <input type="password"
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={confirmPasswordInput} onChange={(e) => setConfirmPasswordInput(e.target.value)} />
              </div>
            </div>
            <button onClick={handleChangePassword}
              className="w-full mt-4 bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 shadow-sm shadow-teal-800/30 ring-1 ring-inset ring-white/10 transition-colors">
              Update password
            </button>
          </div>
        )}

        {/* Summary cards */}
        <div className="flex items-center justify-between mb-2">
          <p className="text-sm text-stone-500">
            Totals for: <span className="font-semibold text-stone-700">
              {selectedYear.length ? selectedYear.join(", ") : ""}
              {selectedMonth.length ? ` · ${selectedMonth.map(monthLabel).join(", ")}` : ""}
              {selectedCompany.length ? ` · ${selectedCompany.join(", ")}` : ""}
              {selectedEmployee.length ? ` · ${selectedEmployee.join(", ")}` : ""}
              {selectedSupplier.length ? ` · ${selectedSupplier.join(", ")}` : ""}
              {!hasActiveFilter && "all months"}
            </span>
          </p>
        </div>
        <div className="grid grid-cols-3 gap-1.5 sm:gap-3 mb-6">
          <div className="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0">
            <div className="bg-stone-100 rounded-xl p-1.5 sm:p-2 text-stone-600 shrink-0"><Ticket size={18} className="sm:hidden" /><Ticket size={20} className="hidden sm:block" /></div>
            <div className="min-w-0">
              <p className="text-xs text-stone-500">Tickets</p>
              <p className="text-sm sm:text-lg font-bold truncate">{totals.count}</p>
            </div>
          </div>
          <div className="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0">
            <div className="bg-teal-50 rounded-xl p-1.5 sm:p-2 text-teal-900 shrink-0"><Wallet size={18} className="sm:hidden" /><Wallet size={20} className="hidden sm:block" /></div>
            <div className="min-w-0">
              <p className="text-xs text-stone-500">Total sales</p>
              <p className="text-sm sm:text-lg font-bold truncate">{fmt(totals.total)}</p>
            </div>
          </div>
          <div className="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0">
            <div className="bg-emerald-50 rounded-xl p-1.5 sm:p-2 text-emerald-700 shrink-0"><TrendingUp size={18} className="sm:hidden" /><TrendingUp size={20} className="hidden sm:block" /></div>
            <div className="min-w-0">
              <p className="text-xs text-stone-500">Total profit</p>
              <p className="text-sm sm:text-lg font-bold text-emerald-700 truncate">{fmt(totals.profit)}</p>
            </div>
          </div>
        </div>

        {/* Entry form: hidden for accounting accounts (view-only + notes-only), and for
            anyone with neither add nor edit permission. Shown while editing an existing
            ticket as long as the person has edit access, even if add access is off. */}
        {!isAccountingUser && (canAddTickets || (form.id && canEditTickets)) && (
        <div className="bg-white rounded-2xl border border-stone-200 p-4 md:p-5 mb-6">
          <h2 className="font-semibold text-stone-900 mb-4">{form.id ? "Edit ticket" : "Add a new ticket"}</h2>
          {error && (
            <div className="bg-red-50 text-red-700 text-sm rounded-xl px-3 py-2 mb-3">{error}</div>
          )}
          <div className="max-w-xs">
            <label className="text-xs text-stone-500 block mb-1">Entered by</label>
            <div className="w-full border border-stone-200 bg-stone-50 rounded-xl px-3 py-2 text-sm text-stone-600">
              {currentUser.name}
            </div>
          </div>

          {/* Reissue / Refund: a single box where you pick which one applies to this
              ticket, instead of two separate checkbox boxes. Picking one clears/closes
              the other. */}
          <div className="mt-4 bg-stone-50 border border-stone-200 rounded-xl p-3">
            <p className="text-xs font-semibold text-stone-500 mb-2">This ticket is...</p>
            <div className="flex flex-wrap gap-4 text-sm mb-1">
              <label className="flex items-center gap-1.5 cursor-pointer select-none text-stone-700">
                <input
                  type="radio"
                  name="ticketSpecialType"
                  className="w-4 h-4 accent-stone-600"
                  checked={!form.isReissued && !refundBoxOpen}
                  onChange={() => {
                    setForm({ ...form, isReissued: false, oldTicketNumber: "", oldTicketIssueDate: "" });
                    clearAllRefundRows();
                    setRefundBoxOpen(false);
                    setRefundRows([{ number: "", airlineAmount: "", customerAmount: "", customerIndex: 0 }]);
                    setRefundSaved(false);
                  }}
                />
                New ticket
              </label>
              <label className="flex items-center gap-1.5 cursor-pointer select-none text-amber-800">
                <input
                  type="radio"
                  name="ticketSpecialType"
                  className="w-4 h-4 accent-amber-700"
                  checked={form.isReissued}
                  onChange={() => {
                    setForm({ ...form, isReissued: true });
                    clearAllRefundRows();
                    setRefundBoxOpen(false);
                    setRefundRows([{ number: "", airlineAmount: "", customerAmount: "", customerIndex: 0 }]);
                    setRefundSaved(false);
                  }}
                />
                Exchange Ticket
              </label>
              <label className="flex items-center gap-1.5 cursor-pointer select-none text-sky-800">
                <input
                  type="radio"
                  name="ticketSpecialType"
                  className="w-4 h-4 accent-sky-700"
                  checked={refundBoxOpen}
                  onChange={() => {
                    setForm({ ...form, isReissued: false, oldTicketNumber: "", oldTicketIssueDate: "" });
                    setRefundBoxOpen(true);
                  }}
                />
                Refund Ticket
              </label>
            </div>

            {form.isReissued && (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mt-3">
                <div>
                  <label className="text-xs text-stone-500 block mb-1">Old ticket number</label>
                  <input
                    className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                    value={form.oldTicketNumber}
                    onChange={(e) => handleOldTicketNumberChange(e.target.value)}
                    onBlur={handleOldTicketNumberBlur}
                    placeholder="e.g. 077-1234567890"
                  />
                </div>
                <div>
                  <label className="text-xs text-stone-500 block mb-1">Old ticket issue date</label>
                  <div className="w-full border border-stone-200 bg-stone-50 rounded-xl px-3 py-2 text-sm text-stone-600">
                    {form.oldTicketIssueDate
                      ? formatDisplayDate(form.oldTicketIssueDate)
                      : form.oldTicketNumber
                      ? "Not found among saved tickets"
                      : "Enter the old ticket number above"}
                  </div>
                </div>
              </div>
            )}

            {refundBoxOpen && (
              <div className="mt-3 space-y-3">
                {refundRows.map((row, index) => {
                  const target = findTicketByNumber(row.number);
                  const targetCustomers = target ? getCustomers(target) : [];
                  return (
                    <div key={index} className="bg-white border border-sky-200 rounded-xl p-3">
                      <div className="flex items-start gap-2">
                        <div className="flex-1">
                          <label className="text-xs text-stone-500 block mb-1">
                            Ticket number to refund {refundRows.length > 1 ? `#${index + 1}` : ""}
                          </label>
                          <input
                            className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                            value={row.number}
                            onChange={(e) => handleRefundRowNumberChange(index, e.target.value)}
                            onBlur={() => handleRefundRowNumberBlur(index)}
                            placeholder="e.g. 077-1234567890"
                          />
                        </div>
                        {refundRows.length > 1 && (
                          <button
                            type="button"
                            onClick={() => removeRefundRow(index)}
                            className="mt-6 text-stone-400 hover:text-red-600"
                            title="Remove this ticket"
                          >
                            <X size={16} />
                          </button>
                        )}
                      </div>

                      {!target ? (
                        <p className="text-xs text-stone-400 mt-2">
                          {row.number ? "Not found among saved tickets" : "Enter the ticket number above"}
                        </p>
                      ) : (
                        <div className="mt-3">
                          <div className="bg-sky-50 border border-sky-200 rounded-xl px-3 py-2 text-sm mb-3">
                            <p className="text-xs text-sky-500 mb-1">Ticket found</p>
                            <p className="text-sky-900 font-medium">{routeLabel(target)}</p>
                            <p className="text-stone-600 text-xs mt-1">
                              {targetCustomers.map((c) => c.name || "-").join(", ")} · {fmt(target.soldPrice)}
                            </p>
                          </div>
                          {targetCustomers.length > 1 && (
                            <div className="mb-3">
                              <label className="text-xs text-stone-500 block mb-1">Refunded ticket</label>
                              <select
                                className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white"
                                value={row.customerIndex}
                                onChange={(e) => {
                                  const newIndex = Number(e.target.value);
                                  const existing = getRefunds(target).find((r) => (r.customerIndex || 0) === newIndex);
                                  setRefundRows(
                                    refundRows.map((r, i) =>
                                      i === index
                                        ? {
                                            ...r,
                                            customerIndex: newIndex,
                                            airlineAmount: existing ? existing.airlineAmount || "" : "",
                                            customerAmount: existing ? existing.customerAmount || "" : "",
                                          }
                                        : r
                                    )
                                  );
                                }}
                              >
                                {targetCustomers.map((c, i) => (
                                  <option key={i} value={i}>
                                    {(c.name || `Customer ${i + 1}`) + (c.ticketNumber ? ` — ${c.ticketNumber}` : "")}
                                  </option>
                                ))}
                              </select>
                            </div>
                          )}
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                            <div>
                              <label className="text-xs text-stone-500 block mb-1">Refunded by airline</label>
                              <input
                                type="number"
                                className="w-28 border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 price-input"
                                value={row.airlineAmount}
                                onChange={(e) =>
                                  setRefundRows(refundRows.map((r, i) => (i === index ? { ...r, airlineAmount: e.target.value } : r)))
                                }
                                placeholder="0"
                              />
                            </div>
                            <div>
                              <label className="text-xs text-stone-500 block mb-1">Refunded to customer</label>
                              <input
                                type="number"
                                className="w-28 border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 price-input"
                                value={row.customerAmount}
                                onChange={(e) =>
                                  setRefundRows(refundRows.map((r, i) => (i === index ? { ...r, customerAmount: e.target.value } : r)))
                                }
                                placeholder="0"
                              />
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}

                <div className="flex items-center gap-3">
                  <button
                    type="button"
                    onClick={addRefundRow}
                    className="text-xs font-semibold text-sky-700 hover:text-sky-900"
                  >
                    + Add another ticket
                  </button>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    onClick={saveAllRefunds}
                    className="bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 flex items-center gap-1.5 shadow-sm shadow-teal-800/30 ring-1 ring-inset ring-white/10"
                  >
                    <Check size={15} /> Save refund{refundRows.length > 1 ? "s" : ""}
                  </button>
                  {refundSaved && (
                    <span className="text-xs text-emerald-700 font-medium">Saved</span>
                  )}
                </div>

              </div>
            )}
          </div>

          <div className="flex flex-wrap items-start gap-2 mt-4">
            <div className="flex-1 min-w-[160px]">
              <label className="text-xs text-stone-500 block mb-1">Company (optional)</label>
              <input
                className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={form.company}
                onChange={(e) => setForm({ ...form, company: e.target.value.toUpperCase() })}
                placeholder="e.g. Acme Corp"
                list="company-suggestions"
              />
            </div>
            <div className="w-40 shrink-0">
              <label className="text-xs text-stone-500 block mb-1">Supplier</label>
              {supplierOther ? (
                <div className="flex gap-2">
                  <input
                    className={`w-full border rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 ${form.supplier.trim() ? "border-blue-400 text-blue-700 font-medium bg-blue-50" : "border-stone-300"}`}
                    value={form.supplier}
                    onChange={(e) => setForm({ ...form, supplier: e.target.value })}
                    placeholder="Enter supplier name"
                    autoFocus
                  />
                  <button
                    type="button"
                    onClick={() => { setSupplierOther(false); setForm({ ...form, supplier: "" }); }}
                    className="shrink-0 text-xs text-stone-500 hover:text-teal-800 border border-stone-300 rounded-xl px-2"
                  >
                    List
                  </button>
                </div>
              ) : (
                <select
                  className={`w-full border rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 ${form.supplier ? "border-blue-400 text-blue-700 font-medium bg-blue-50" : "border-stone-300 bg-white"}`}
                  value={form.supplier}
                  onChange={(e) => {
                    if (e.target.value === "__other__") {
                      setSupplierOther(true);
                      setForm({ ...form, supplier: "" });
                    } else {
                      setForm({ ...form, supplier: e.target.value });
                    }
                  }}
                >
                  <option value="">Select supplier</option>
                  {SUPPLIERS.map((s) => (
                    <option key={s} value={s}>{s}</option>
                  ))}
                  <option value="__other__">Other</option>
                </select>
              )}
            </div>
            <div className="w-14 shrink-0">
              <label className="text-xs text-stone-500 block mb-1">Customers</label>
              <input
                type="number"
                min={1}
                max={50}
                className="w-14 border border-stone-300 rounded-xl px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={form.customersCount}
                onChange={(e) => handleCustomersCountChange(e.target.value)}
                onBlur={(e) => {
                  if (e.target.value === "" || parseInt(e.target.value, 10) < 1) {
                    handleCustomersCountChange(1);
                  }
                }}
                placeholder="1"
              />
            </div>
          </div>

          {/* Dynamic customer name + ticket number cells, one row per customer. A
              "Conjunction" checkbox sits between the name and ticket number — check it
              when that customer has a second ticket number issued together with the
              first, which reveals a second field for its "-XXX" suffix inside the same
              ticket number box. */}
          <div className="mt-4">
            <label className="text-xs text-stone-500 block mb-2">
              Customers ({form.customers.length})
            </label>
            <div className="space-y-2">
              {form.customers.map((c, i) => (
                <div key={i} className="flex flex-col md:flex-row gap-2 md:gap-3 md:items-start">
                  <input
                    className="w-full md:flex-1 border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                    value={c.name}
                    onChange={(e) => handleCustomerFieldChange(i, "name", e.target.value)}
                    placeholder={`Customer ${i + 1} name`}
                  />
                  <label
                    className="flex items-center gap-1.5 shrink-0 cursor-pointer select-none text-xs text-stone-500 md:py-2"
                    title="This customer has a second ticket number issued together with the first"
                  >
                    <input
                      type="checkbox"
                      className="w-4 h-4 accent-stone-600"
                      checked={!!c.conjunction}
                      onChange={(e) => handleCustomerConjunctionToggle(i, e.target.checked)}
                    />
                    Conjunction
                  </label>
                  <div className="w-full md:w-[24ch] md:shrink-0 flex items-center border border-stone-300 rounded-xl px-3 py-2 focus-within:ring-2 focus-within:ring-teal-700">
                    <input
                      className="min-w-0 text-sm outline-none bg-transparent flex-1"
                      style={c.conjunction ? { flex: "0 0 auto", width: `${Math.max((c.ticketNumber || "").length - ((c.ticketNumber || "").match(/-/g) || []).length * 0.5, 3)}ch` } : { width: "20ch" }}
                      value={c.ticketNumber}
                      onChange={(e) => handleCustomerFieldChange(i, "ticketNumber", e.target.value)}
                      onBlur={() => handleTicketNumberBlur(i)}
                      placeholder={`Ticket number ${i + 1}`}
                    />
                    {c.conjunction && (c.ticketNumber || "").replace(/[^A-Z0-9]/g, "").length >= 13 && (
                      <>
                        <span className="text-stone-800 font-semibold mx-0.5 select-none">-</span>
                        <input
                          className="min-w-0 text-sm outline-none bg-transparent"
                          style={{ flex: "0 0 auto", width: `${Math.max((c.ticketNumber2 || "").replace(/^-/, "").length, 1) + 1}ch` }}
                          value={(c.ticketNumber2 || "").replace(/^-/, "")}
                          onChange={(e) => handleCustomerFieldChange(i, "ticketNumber2", `-${e.target.value.replace(/^-/, "")}`)}
                          placeholder="891"
                        />
                      </>
                    )}
                  </div>
                  <input
                    className="w-full md:w-[13ch] md:shrink-0 border border-stone-300 rounded-xl px-3 py-2 text-sm font-mono uppercase outline-none focus:ring-2 focus:ring-teal-700"
                    value={c.pnrReference || ""}
                    onChange={(e) => handleCustomerFieldChange(i, "pnrReference", e.target.value)}
                    onBlur={() => handlePnrReferenceBlur(i)}
                    placeholder="PNR ref"
                    maxLength={6}
                    title="Booking PNR reference (up to 6 letters/digits)"
                  />
                </div>
              ))}
            </div>
          </div>

          <div className="mt-4 flex flex-wrap items-center gap-x-6 gap-y-2">
            <div className="flex items-center gap-4 text-xs text-stone-500">
              <label className="flex items-center gap-1.5 cursor-pointer select-none">
                <input
                  type="radio"
                  name="routeMode"
                  className="w-4 h-4 accent-teal-800"
                  checked={!form.multiDestination && (form.tripType || "oneWay") === "oneWay"}
                  onChange={() => setForm({ ...form, tripType: "oneWay", multiDestination: false })}
                />
                One way
              </label>
              <label className="flex items-center gap-1.5 cursor-pointer select-none">
                <input
                  type="radio"
                  name="routeMode"
                  className="w-4 h-4 accent-teal-800"
                  checked={!form.multiDestination && form.tripType === "roundTrip"}
                  onChange={() => setForm({ ...form, tripType: "roundTrip", multiDestination: false })}
                />
                Round trip
              </label>
              <label className="flex items-center gap-2 text-xs font-semibold text-stone-500 cursor-pointer select-none">
                <input
                  type="radio"
                  name="routeMode"
                  className="w-4 h-4 accent-teal-800"
                  checked={!!form.multiDestination}
                  onChange={() => {
                    setForm({
                      ...form,
                      multiDestination: true,
                      // Seed the stop list from the current From/To the first time this is
                      // switched on, so nothing already typed gets lost.
                      destinations:
                        !(form.destinations || []).some((d) => (d || "").trim())
                          ? [form.from || "", form.to || ""]
                          : form.destinations,
                    });
                  }}
                />
                Multi-destination route (multi-city)
              </label>
            </div>
          </div>

          <div className="flex flex-wrap items-end gap-2 mt-2">
            {form.multiDestination ? (
              <>
                {form.destinations.map((d, i) => (
                  <div key={i} className="flex items-end gap-1">
                    <div>
                      <label className="text-[10px] text-stone-400 block mb-1">
                        {i === 0 ? "From" : i === form.destinations.length - 1 ? "Final" : `Stop ${i}`}
                      </label>
                      <input
                        className="w-16 border border-stone-300 rounded-lg px-2 py-1.5 text-xs focus:outline-none focus:ring-2 focus:ring-teal-700 uppercase"
                        value={d}
                        onChange={(e) => handleDestinationChange(i, e.target.value)}
                        placeholder={i === 0 ? "CAI" : "DXB"}
                        list="city-suggestions"
                      />
                    </div>
                    {form.destinations.length > 2 && (
                      <button
                        type="button"
                        onClick={() => removeDestinationStop(i)}
                        className="shrink-0 text-stone-400 hover:text-red-600 mb-1.5"
                        title="Remove stop"
                      >
                        <X size={14} />
                      </button>
                    )}
                  </div>
                ))}
                <button
                  type="button"
                  onClick={addDestinationStop}
                  className="inline-flex items-center gap-1 text-xs font-semibold text-teal-700 hover:text-teal-900 mb-1.5"
                >
                  <Plus size={14} /> Add stop
                </button>
              </>
            ) : (
              <>
                <div>
                  <label className="text-xs text-stone-500 block mb-1">From</label>
                  <input
                    className="w-16 border border-stone-300 rounded-lg px-2 py-1.5 text-xs focus:outline-none focus:ring-2 focus:ring-teal-700 uppercase"
                    value={form.from}
                    onChange={(e) => handleCityChange("from", e.target.value)}
                    placeholder="CAI"
                    list="city-suggestions"
                  />
                </div>
                <div>
                  <label className="text-xs text-stone-500 block mb-1">To</label>
                  <input
                    className="w-16 border border-stone-300 rounded-lg px-2 py-1.5 text-xs focus:outline-none focus:ring-2 focus:ring-teal-700 uppercase"
                    value={form.to}
                    onChange={(e) => handleCityChange("to", e.target.value)}
                    placeholder="DXB"
                    list="city-suggestions"
                  />
                </div>
                {form.tripType === "roundTrip" && (
                  <div>
                    <label className="text-xs text-stone-500 block mb-1">Return airport</label>
                    <div
                      className="w-16 border border-stone-200 bg-stone-50 rounded-lg px-2 py-1.5 text-xs text-stone-600 uppercase truncate"
                      title="Automatically matches the first (From) airport"
                    >
                      {form.from || "-"}
                    </div>
                  </div>
                )}
              </>
            )}
            <div>
              <label className="text-xs text-stone-500 mb-1 flex items-center gap-1.5">
                <span>Airline</span>
                {getAirlineNameByIata(form.airline) && (
                  <span className="bg-teal-50 text-teal-700 border border-teal-200 rounded px-1.5 py-0.5 text-[10px] font-semibold">
                    {getAirlineNameByIata(form.airline)}
                  </span>
                )}
              </label>
              <input
                className="w-16 border border-stone-300 rounded-lg px-2 py-1.5 text-xs focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={form.airline}
                onChange={(e) => handleAirlineChange(e.target.value)}
                placeholder="MS"
                list="airline-suggestions"
              />
            </div>
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-3">
            <div>
              <label className="text-xs text-stone-500 block mb-1">Ticket issue date</label>
              <input
                type="date"
                lang="en-GB"
                max={todayDateStr()}
                className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={form.date}
                onChange={(e) => {
                  const v = e.target.value;
                  // Belt-and-braces: some browsers still let a future date be typed
                  // manually even with `max` set, so clamp it back to today here too.
                  setForm({ ...form, date: v > todayDateStr() ? todayDateStr() : v });
                }}
              />
            </div>
            <div>
              <label className="text-xs text-stone-500 block mb-1">Net price</label>
              <input
                type="number"
                className="w-28 border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 price-input"
                value={form.netPrice}
                onChange={(e) => setForm({ ...form, netPrice: e.target.value })}
                placeholder="0"
              />
            </div>
            <div>
              <label className="text-xs text-stone-500 block mb-1">Sold price</label>
              <input
                type="number"
                className="w-28 border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 price-input"
                value={form.soldPrice}
                onChange={(e) => setForm({ ...form, soldPrice: e.target.value })}
                placeholder="0"
              />
            </div>
            <div>
              <label className="text-xs text-stone-500 block mb-1">Profit (auto)</label>
              <div className="w-full border border-stone-200 bg-stone-50 rounded-xl px-3 py-2 text-sm text-emerald-700 font-semibold">
                {fmt(profit(form.netPrice, form.soldPrice))}
              </div>
            </div>
            <div className="col-span-4">
              <label className="text-xs text-stone-500 block mb-1">Notes</label>
              <textarea
                className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 min-h-[80px]"
                value={form.notes}
                onChange={(e) => setForm({ ...form, notes: e.target.value.toUpperCase() })}
                placeholder="Optional"
              />
            </div>
          </div>

          <div className="flex gap-2 mt-4">
            <button
              onClick={handleSubmit}
              className="bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 shadow-sm shadow-teal-800/30 ring-1 ring-inset ring-white/10 transition-colors flex items-center gap-1.5"
            >
              <Check size={16} /> {form.id ? "Save changes" : "Add ticket"}
            </button>
            {form.id && (
              <button
                onClick={handleCancel}
                className="border border-stone-300 text-stone-600 text-sm rounded-xl px-4 py-2 flex items-center gap-1.5"
              >
                <X size={16} /> Cancel
              </button>
            )}
          </div>
        </div>
        )}

        {/* IATA balance tracker: the balance itself (editable directly, turns red when
            negative) and a separate box for the value of each newly issued ticket —
            entering a value there and pressing Enter subtracts it from the balance
            above automatically (no separate Deduct button). The History button opens a
            popup listing every amount deducted today — it resets empty at the start of
            each new day. Both fields
            live entirely in their own shared-storage keys (tickets:iataBalance /
            tickets:iataHistory) — they never read from or write into tickets,
            customers, or any other account/total elsewhere in the app. Number spin
            arrows are removed from both via the shared .price-input class. */}
        <div className="bg-white border border-stone-200 rounded-2xl p-3 sm:p-4 mb-3 flex flex-wrap items-end gap-3">
          <div>
            <label className="text-xs text-stone-500 block mb-1">IATA balance</label>
            <div className="relative">
              <Wallet size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none" />
              <input
                type="number"
                className={`price-input w-40 border rounded-xl pl-9 pr-3 py-2 text-sm font-semibold focus:outline-none focus:ring-2 focus:ring-teal-700 ${
                  iataBalance !== null && iataBalance < 0
                    ? "border-red-300 text-red-600 bg-red-50"
                    : "border-stone-300 text-stone-800"
                }`}
                value={iataBalance ?? ""}
                onChange={(e) => setIataBalance(e.target.value === "" ? null : parseFloat(e.target.value))}
                onBlur={() => iataBalance !== null && !Number.isNaN(iataBalance) && persistIataBalance(iataBalance)}
                placeholder="0"
              />
            </div>
          </div>
          <div>
            <label className="text-xs text-stone-500 block mb-1">Issued ticket value</label>
            <div className="relative">
              <Ticket size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400 pointer-events-none" />
              <input
                type="number"
                className="price-input w-40 border border-stone-300 rounded-xl pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={iataTicketValueInput}
                onChange={(e) => setIataTicketValueInput(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && applyIataTicketValue()}
                placeholder="0"
              />
            </div>
          </div>
          <button
            type="button"
            onClick={() => setShowIataHistory(true)}
            className="shrink-0 inline-flex items-center gap-1.5 text-xs font-semibold text-teal-800 border border-teal-800 rounded-xl px-3 py-2 hover:bg-teal-50"
          >
            <History size={14} /> History
          </button>
        </div>

        {/* Search and filters — one unified card: search + a "Filters" toggle with a
            count badge, an optional expanded panel with the dropdowns, and a row of
            removable chips for whatever is currently active. */}
        <div className="bg-white border border-stone-200 rounded-2xl p-3 sm:p-4 mb-3">
          <div className="flex items-stretch gap-2">
            <div className="relative flex-1">
              <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
              <input
                className="w-full border border-stone-300 rounded-xl pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                placeholder="Search by employee, company, ticket number, customer, destination, or airline"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
              />
            </div>
            <button
              type="button"
              onClick={() => setFiltersOpen(!filtersOpen)}
              className={`shrink-0 flex items-center gap-1.5 border rounded-xl px-3 py-2 text-sm font-medium transition-colors ${
                filtersOpen ? "border-teal-700 text-teal-800 bg-teal-50" : "border-stone-300 text-stone-600 hover:bg-stone-50 bg-white"
              }`}
            >
              <SlidersHorizontal size={16} />
              <span className="hidden sm:inline">Filters</span>
              {activeFilterCount > 0 && (
                <span className="bg-teal-700 text-white text-[11px] font-bold rounded-full min-w-[20px] h-5 px-1 flex items-center justify-center">
                  {activeFilterCount}
                </span>
              )}
              <ChevronDown size={14} className={`transition-transform ${filtersOpen ? "rotate-180" : ""}`} />
            </button>
            <button
              type="button"
              onClick={() => { if (hasActiveFilter) exportFiltered(); }}
              disabled={!hasActiveFilter}
              title={hasActiveFilter ? "" : "Select at least one filter (year, month, company, employee, supplier, or search) before exporting"}
              className={`shrink-0 flex items-center gap-1.5 border rounded-xl px-3 py-2 text-sm font-medium transition-colors ${
                hasActiveFilter
                  ? "text-teal-800 border-teal-800 hover:bg-teal-50 bg-white"
                  : "text-stone-400 border-stone-200 cursor-not-allowed bg-white"
              }`}
            >
              <Download size={16} />
              <span className="hidden sm:inline">{hasActiveFilter ? "Export to Excel" : "Select a filter to export"}</span>
            </button>
          </div>

          {filtersOpen && (
            <div className="flex flex-wrap gap-3 mt-3 pt-3 border-t border-stone-100">
              <div>
                <label className="text-xs text-stone-500 block mb-1">Year</label>
                <MultiSelectDropdown
                  label="years"
                  icon={Calendar}
                  options={yearsAvailable}
                  selected={selectedYear}
                  onChange={setSelectedYear}
                  placeholder="All years"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Month</label>
                <MultiSelectDropdown
                  label="months"
                  icon={Calendar}
                  options={monthsAvailable.map((key) => ({ value: key, label: monthLabel(key) }))}
                  selected={selectedMonth}
                  onChange={setSelectedMonth}
                  placeholder="All months"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Company</label>
                <MultiSelectDropdown
                  label="companies"
                  icon={Building2}
                  options={companiesAvailable}
                  selected={selectedCompany}
                  onChange={setSelectedCompany}
                  placeholder="All companies"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Employee</label>
                <MultiSelectDropdown
                  label="employees"
                  icon={User}
                  options={employeesAvailable}
                  selected={selectedEmployee}
                  onChange={setSelectedEmployee}
                  placeholder="All employees"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Supplier</label>
                <MultiSelectDropdown
                  label="suppliers"
                  icon={Plane}
                  options={suppliersAvailable}
                  selected={selectedSupplier}
                  onChange={setSelectedSupplier}
                  placeholder="All suppliers"
                />
              </div>
            </div>
          )}

          {hasActiveFilter && (
            <div className="flex flex-wrap items-center gap-x-4 gap-y-1.5 mt-3 pt-3 border-t border-stone-100 text-xs">
              <span className="text-stone-400 font-medium">Applied:</span>
              {selectedYear.map((y) => (
                <span key={`year-${y}`} className="inline-flex items-center gap-1 text-stone-600">
                  Year: <span className="font-semibold text-stone-800">{y}</span>
                  <button onClick={() => setSelectedYear(selectedYear.filter((v) => v !== y))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              ))}
              {selectedMonth.map((m) => (
                <span key={`month-${m}`} className="inline-flex items-center gap-1 text-stone-600">
                  Month: <span className="font-semibold text-stone-800">{monthLabel(m)}</span>
                  <button onClick={() => setSelectedMonth(selectedMonth.filter((v) => v !== m))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              ))}
              {selectedCompany.map((c) => (
                <span key={`company-${c}`} className="inline-flex items-center gap-1 text-stone-600">
                  Company: <span className="font-semibold text-stone-800">{c}</span>
                  <button onClick={() => setSelectedCompany(selectedCompany.filter((v) => v !== c))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              ))}
              {selectedEmployee.map((e) => (
                <span key={`employee-${e}`} className="inline-flex items-center gap-1 text-stone-600">
                  Employee: <span className="font-semibold text-stone-800">{e}</span>
                  <button onClick={() => setSelectedEmployee(selectedEmployee.filter((v) => v !== e))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              ))}
              {selectedSupplier.map((s) => (
                <span key={`supplier-${s}`} className="inline-flex items-center gap-1 text-stone-600">
                  Supplier: <span className="font-semibold text-stone-800">{s}</span>
                  <button onClick={() => setSelectedSupplier(selectedSupplier.filter((v) => v !== s))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              ))}
              {query.trim() && (
                <span className="inline-flex items-center gap-1 text-stone-600">
                  Search: <span className="font-semibold text-stone-800">"{query.trim()}"</span>
                  <button onClick={() => setQuery("")} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              )}
              <button
                onClick={clearAllFilters}
                className="text-red-600 hover:text-red-700 font-semibold ml-auto"
              >
                Clear all
              </button>
            </div>
          )}
        </div>

        <datalist id="company-suggestions">
          {suggestions.companies.map((c) => (
            <option key={companyName(c)} value={companyName(c)} />
          ))}
        </datalist>
        <datalist id="airline-suggestions">
          {suggestions.airlines.map((code) => (
            <option key={`u-${code}`} value={code} />
          ))}
          {AIRLINE_CODES.map((a) => (
            <option key={`a-${a.code}`} value={a.iata} label={`${a.iata} — ${a.name}`} />
          ))}
        </datalist>
        <datalist id="city-suggestions">
          {suggestions.cities.map((name) => (
            <option key={`u-${name}`} value={name} />
          ))}
          {AIRPORTS.map((entry) => (
            <option key={`p-${entry}`} value={entry} />
          ))}
        </datalist>

        {/* Ticket list */}
        <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden">
          {filtered.length === 0 ? (
            <p className="text-center text-stone-400 text-sm py-10">
              {visibleTickets.length === 0 ? "No tickets recorded yet" : "No results match your search"}
            </p>
          ) : (
            <div className="overflow-x-auto rounded-xl border border-stone-200">
              <table className="w-full text-xs border-collapse">
                <thead>
                  <tr className="bg-teal-50/60 text-teal-800 text-[11px] uppercase tracking-wide border-b-2 border-teal-200">
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Employee</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Date</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Customer</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Ticket #</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Airline</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Route</th>
                    <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Net price</th>
                    <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Sold price</th>
                    <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Profit</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Company</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Supplier</th>
                  </tr>
                </thead>
                <tbody>
                  {sortedFiltered.flatMap((t) => buildTicketRows(t))}
                </tbody>
              </table>
            </div>
          )}
        </div>

        {selectedMonth.length === 0 && monthlyBreakdown.length > 0 && (
          <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden mt-6">
            <div className="px-4 py-3 border-b border-stone-100">
              <h2 className="font-semibold text-stone-900 text-sm">Totals by month</h2>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-stone-50 text-stone-500 text-xs">
                    <th className="text-left px-3 py-2 font-medium">Month</th>
                    <th className="text-left px-3 py-2 font-medium">Tickets</th>
                    <th className="text-left px-3 py-2 font-medium">Total sales</th>
                    <th className="text-left px-3 py-2 font-medium">Total profit</th>
                    <th className="text-left px-3 py-2 font-medium"></th>
                  </tr>
                </thead>
                <tbody>
                  {monthlyBreakdown.map((m) => (
                    <tr key={m.key} className="border-t border-stone-100 hover:bg-stone-50">
                      <td className="px-3 py-2 font-medium text-stone-800">{monthLabel(m.key)}</td>
                      <td className="px-3 py-2 text-stone-600">{m.count}</td>
                      <td className="px-3 py-2 text-stone-600">{fmt(m.total)}</td>
                      <td className="px-3 py-2 font-semibold text-emerald-700">{fmt(m.profit)}</td>
                      <td className="px-3 py-2">
                        <div className="flex items-center gap-3 justify-end">
                          <button
                            onClick={() => exportMonth(m.key)}
                            className="text-stone-400 hover:text-teal-800 text-xs font-medium flex items-center gap-1"
                          >
                            <Download size={13} /> Export
                          </button>
                          <button
                            onClick={() => setSelectedMonth(m.key)}
                            className="text-teal-800 text-xs font-medium hover:underline"
                          >
                            View details
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {selectedCompany.length === 0 && companyBreakdown.length > 0 && (
          <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden mt-6">
            <div className="px-4 py-3 border-b border-stone-100">
              <h2 className="font-semibold text-stone-900 text-sm">Companies and their customers</h2>
            </div>
            <div className="divide-y divide-stone-100">
              {companyBreakdown.map((c) => (
                <div key={c.name} className="px-4 py-3">
                  <div className="flex items-center justify-between flex-wrap gap-2">
                    <div className="flex items-center gap-2">
                      <Building2 size={16} className="text-stone-400" />
                      <button
                        onClick={() => setSelectedCompany(c.name)}
                        className="font-medium text-stone-800 hover:text-teal-800 hover:underline text-sm"
                      >
                        {c.name}
                      </button>
                      <span className="text-xs text-stone-400">({c.count} tickets)</span>
                    </div>
                    <div className="flex items-center gap-4 text-xs text-stone-500">
                      <span>Sales: <span className="font-semibold text-stone-700">{fmt(c.total)}</span></span>
                      <span>Profit: <span className="font-semibold text-emerald-700">{fmt(c.profit)}</span></span>
                    </div>
                  </div>
                  <p className="text-xs text-stone-500 mt-1.5 pl-6">
                    Customers: {c.customers.join(", ")}
                  </p>
                </div>
              ))}
            </div>
          </div>
        )}

        <p className="text-xs text-stone-400 mt-3">
          This data is shared between signed-in employees. Login is a basic access gate, not strong security — treat it accordingly.
        </p>
        </>
        )}

        {activeSection === "hotels" && (
        <>
        {/* Summary cards, same style as the Flights section */}
        <div className="grid grid-cols-3 gap-1.5 sm:gap-3 mb-6">
          <div className="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0">
            <div className="bg-stone-100 rounded-xl p-1.5 sm:p-2 text-stone-600 shrink-0"><Building2 size={18} className="sm:hidden" /><Building2 size={20} className="hidden sm:block" /></div>
            <div className="min-w-0">
              <p className="text-xs text-stone-500">Bookings</p>
              <p className="text-sm sm:text-lg font-bold truncate">{hotelTotals.count}</p>
            </div>
          </div>
          <div className="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0">
            <div className="bg-teal-50 rounded-xl p-1.5 sm:p-2 text-teal-900 shrink-0"><Wallet size={18} className="sm:hidden" /><Wallet size={20} className="hidden sm:block" /></div>
            <div className="min-w-0">
              <p className="text-xs text-stone-500">Total sales (EGP)</p>
              <p className="text-sm sm:text-lg font-bold truncate">{fmt(hotelTotals.sold)}</p>
            </div>
          </div>
          <div className="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0">
            <div className="bg-emerald-50 rounded-xl p-1.5 sm:p-2 text-emerald-700 shrink-0"><TrendingUp size={18} className="sm:hidden" /><TrendingUp size={20} className="hidden sm:block" /></div>
            <div className="min-w-0">
              <p className="text-xs text-stone-500">Total profit (EGP)</p>
              <p className="text-sm sm:text-lg font-bold text-emerald-700 truncate">{fmt(hotelTotals.profit)}</p>
            </div>
          </div>
        </div>

        {/* Buttons to register new supplier names and hotel names, so they're always
            available to pick from the Supplier / Hotel name fields below. */}
        <div className="flex flex-wrap items-center gap-2 mb-4">
          <button
            onClick={() => { setShowAddSupplierPanel(!showAddSupplierPanel); setShowAddHotelNamePanel(false); }}
            className="text-xs font-semibold text-teal-800 border border-teal-700 rounded-xl px-3 py-2 hover:bg-teal-50 flex items-center gap-1.5"
          >
            <Plus size={14} /> Add supplier
          </button>
          <button
            onClick={() => { setShowAddHotelNamePanel(!showAddHotelNamePanel); setShowAddSupplierPanel(false); }}
            className="text-xs font-semibold text-teal-800 border border-teal-700 rounded-xl px-3 py-2 hover:bg-teal-50 flex items-center gap-1.5"
          >
            <Plus size={14} /> Add hotel name
          </button>
        </div>

        {showAddSupplierPanel && (
          <div className="bg-white border border-stone-200 rounded-2xl p-4 mb-4">
            <h3 className="text-sm font-bold text-stone-700 mb-3">Suppliers</h3>
            <div className="flex gap-2 mb-3">
              <input
                className="w-full max-w-xs border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={newSupplierDraft}
                onChange={(e) => setNewSupplierDraft(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleAddSupplierName()}
                placeholder="Supplier name"
              />
              <button
                onClick={handleAddSupplierName}
                className="bg-gradient-to-b from-teal-700 to-teal-900 text-white text-sm font-semibold rounded-xl px-4 py-2 hover:brightness-110"
              >
                Add
              </button>
            </div>
            {suggestions.suppliers.length === 0 ? (
              <p className="text-xs text-stone-400">No suppliers saved yet</p>
            ) : (
              <div className="flex flex-wrap gap-2">
                {suggestions.suppliers.map((s) => (
                  <span
                    key={s}
                    className="inline-flex items-center gap-1.5 bg-stone-50 border border-stone-200 rounded-lg px-2.5 py-1 text-xs text-stone-700"
                  >
                    {s}
                    <button onClick={() => handleDeleteSupplierName(s)} className="text-red-500 hover:text-red-700">
                      <Trash2 size={12} />
                    </button>
                  </span>
                ))}
              </div>
            )}
          </div>
        )}

        {showAddHotelNamePanel && (
          <div className="bg-white border border-stone-200 rounded-2xl p-4 mb-4">
            <h3 className="text-sm font-bold text-stone-700 mb-3">Hotel names</h3>
            <div className="flex gap-2 mb-3">
              <input
                className="w-full max-w-xs border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={newHotelNameDraft}
                onChange={(e) => setNewHotelNameDraft(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleAddHotelName()}
                placeholder="Hotel name"
              />
              <button
                onClick={handleAddHotelName}
                className="bg-gradient-to-b from-teal-700 to-teal-900 text-white text-sm font-semibold rounded-xl px-4 py-2 hover:brightness-110"
              >
                Add
              </button>
            </div>
            {suggestions.hotelNames.length === 0 ? (
              <p className="text-xs text-stone-400">No hotel names saved yet</p>
            ) : (
              <div className="flex flex-wrap gap-2">
                {suggestions.hotelNames.map((hn) => (
                  <span
                    key={hn}
                    className="inline-flex items-center gap-1.5 bg-stone-50 border border-stone-200 rounded-lg px-2.5 py-1 text-xs text-stone-700"
                  >
                    {hn}
                    <button onClick={() => handleDeleteHotelName(hn)} className="text-red-500 hover:text-red-700">
                      <Trash2 size={12} />
                    </button>
                  </span>
                ))}
              </div>
            )}
          </div>
        )}

        {/* USD -> EGP exchange rate bar — entered by hand each day (e.g. from the CBE's
            published rate), saved to shared storage so every employee sees the same value. */}
        <div className="bg-white border border-stone-200 rounded-2xl px-4 py-3 mb-4 flex flex-wrap items-center gap-3">
          <span className="text-xs font-semibold text-stone-500">USD → EGP rate today:</span>
          <input
            type="number"
            step="0.01"
            className="w-28 border border-stone-300 rounded-lg px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
            value={usdToEgpRate ?? ""}
            onChange={(e) => setUsdToEgpRate(e.target.value === "" ? null : parseFloat(e.target.value))}
            onBlur={() => {
              if (usdToEgpRate !== null && !Number.isNaN(usdToEgpRate)) persistUsdRate(usdToEgpRate);
            }}
            placeholder="e.g. 51.20"
          />
          <button
            onClick={() => usdToEgpRate !== null && !Number.isNaN(usdToEgpRate) && persistUsdRate(usdToEgpRate)}
            className="text-xs font-semibold text-teal-800 border border-teal-700 rounded-lg px-3 py-1.5 hover:bg-teal-50"
          >
            Save rate
          </button>
          {usdToEgpRateDate && (
            <span className="text-xs text-stone-400">Last updated: {formatDisplayDate(usdToEgpRateDate)}</span>
          )}
        </div>

        {hotelError && (
          <div className="text-sm rounded-xl px-3 py-2 mb-4 bg-red-50 text-red-700">{hotelError}</div>
        )}

        {canAddTickets && (
          <div className="bg-white border border-stone-200 rounded-2xl p-5 mb-6">
            <h3 className="text-sm font-bold text-stone-700 mb-4">
              {hotelEditingId ? "Edit hotel booking" : "New hotel booking"}
            </h3>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
              <div>
                <label className="text-xs text-stone-500 block mb-1">
                  Company name <span className="font-normal text-stone-400">(optional — leave blank for Individual)</span>
                </label>
                <input
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={hotelForm.customer}
                  onChange={(e) => setHotelForm({ ...hotelForm, customer: e.target.value })}
                  placeholder="e.g. Perla Travel Corp — leave blank for Individual"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Hotel name</label>
                {hotelNameOther ? (
                  <div className="flex gap-2">
                    <input
                      className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                      value={hotelForm.hotel}
                      onChange={(e) => setHotelForm({ ...hotelForm, hotel: e.target.value })}
                      placeholder="Enter hotel name"
                      autoFocus
                    />
                    <button
                      type="button"
                      onClick={() => { setHotelNameOther(false); setHotelForm({ ...hotelForm, hotel: "" }); }}
                      className="shrink-0 text-xs text-stone-500 hover:text-teal-800 border border-stone-300 rounded-xl px-2"
                    >
                      List
                    </button>
                  </div>
                ) : (
                  <select
                    className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white"
                    value={hotelForm.hotel}
                    onChange={(e) => {
                      if (e.target.value === "__other__") {
                        setHotelNameOther(true);
                        setHotelForm({ ...hotelForm, hotel: "" });
                      } else {
                        setHotelForm({ ...hotelForm, hotel: e.target.value });
                      }
                    }}
                  >
                    <option value="">Select hotel</option>
                    {suggestions.hotelNames.map((hn) => (
                      <option key={hn} value={hn}>{hn}</option>
                    ))}
                    <option value="__other__">Other</option>
                  </select>
                )}
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Supplier</label>
                {hotelSupplierOther ? (
                  <div className="flex gap-2">
                    <input
                      className={`w-full border rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 ${hotelForm.supplier.trim() ? "border-blue-400 text-blue-700 font-medium bg-blue-50" : "border-stone-300"}`}
                      value={hotelForm.supplier}
                      onChange={(e) => setHotelForm({ ...hotelForm, supplier: e.target.value })}
                      placeholder="Enter supplier name"
                      autoFocus
                    />
                    <button
                      type="button"
                      onClick={() => { setHotelSupplierOther(false); setHotelForm({ ...hotelForm, supplier: "" }); }}
                      className="shrink-0 text-xs text-stone-500 hover:text-teal-800 border border-stone-300 rounded-xl px-2"
                    >
                      List
                    </button>
                  </div>
                ) : (
                  <select
                    className={`w-full border rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 ${hotelForm.supplier ? "border-blue-400 text-blue-700 font-medium bg-blue-50" : "border-stone-300 bg-white"}`}
                    value={hotelForm.supplier}
                    onChange={(e) => {
                      if (e.target.value === "__other__") {
                        setHotelSupplierOther(true);
                        setHotelForm({ ...hotelForm, supplier: "" });
                      } else {
                        setHotelForm({ ...hotelForm, supplier: e.target.value });
                      }
                    }}
                  >
                    <option value="">Select supplier</option>
                    {suggestions.suppliers.map((s) => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                    <option value="__other__">Other</option>
                  </select>
                )}
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Booking date</label>
                <input
                  type="date"
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={hotelForm.bookingDate}
                  onChange={(e) => setHotelForm({ ...hotelForm, bookingDate: e.target.value })}
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Notes</label>
                <input
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={hotelForm.notes}
                  onChange={(e) => setHotelForm({ ...hotelForm, notes: e.target.value })}
                />
              </div>
            </div>

            <p className="text-xs text-stone-500 mb-3">
              Each room has its own check-in/check-out dates — price is per room, per night.
            </p>

            {/* Room lines: one booking can mix different room types, meal plans, currencies,
                prices, and stay dates — each room keeps its own check-in/check-out. */}
            <div className="space-y-3">
              <label className="text-xs text-stone-500 block">Rooms</label>
              {hotelForm.roomLines.map((line) => (
                <div key={line.id} className="bg-stone-50 border border-stone-200 rounded-xl p-3 space-y-3">
                  {/* Row 1: room type, meal plan, dates, currency. */}
                  <div className="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-5 gap-3 items-end">
                    <div>
                      <label className="text-[11px] text-stone-500 block mb-1">Room type</label>
                      <select
                        className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={line.roomType}
                        onChange={(e) => {
                          const roomType = e.target.value;
                          const capacity = ROOM_CAPACITY[roomType] || 1;
                          updateHotelRoomLine(line.id, { roomType, guests: guestsForCapacity(line.guests, capacity) });
                        }}
                      >
                        {ROOM_TYPES.map((r) => (
                          <option key={r.value} value={r.value}>{r.label}</option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="text-[11px] text-stone-500 block mb-1">Meal plan</label>
                      <select
                        className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={line.mealPlan}
                        onChange={(e) => updateHotelRoomLine(line.id, { mealPlan: e.target.value })}
                      >
                        {MEAL_PLANS.map((m) => (
                          <option key={m.value} value={m.value}>{m.label}</option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="text-[11px] text-stone-500 block mb-1">Check-in</label>
                      <input
                        type="date"
                        className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={line.checkIn}
                        onChange={(e) => updateHotelRoomLine(line.id, { checkIn: e.target.value })}
                      />
                    </div>
                    <div>
                      <label className="text-[11px] text-stone-500 block mb-1">Check-out</label>
                      <input
                        type="date"
                        min={line.checkIn || undefined}
                        className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={line.checkOut}
                        onChange={(e) => updateHotelRoomLine(line.id, { checkOut: e.target.value })}
                      />
                    </div>
                    <div>
                      <label className="text-[11px] text-stone-500 block mb-1">Currency</label>
                      <select
                        className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={line.currency}
                        onChange={(e) => updateHotelRoomLine(line.id, { currency: e.target.value })}
                      >
                        {HOTEL_CURRENCIES.map((c) => (
                          <option key={c.value} value={c.value}>{c.label}</option>
                        ))}
                      </select>
                    </div>
                  </div>

                  {/* Row 2: # rooms, net, sold. */}
                  <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 items-start">
                    <div>
                      <label className="text-[11px] text-stone-500 block mb-1"># rooms</label>
                      <input
                        type="number"
                        min="1"
                        className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={line.count}
                        onChange={(e) => updateHotelRoomLine(line.id, { count: e.target.value })}
                      />
                    </div>
                    <div>
                      <label className="text-[11px] text-stone-500 block mb-1">Net (per room/night)</label>
                      <input
                        type="number"
                        className="w-28 border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 price-input"
                        value={line.netPrice}
                        onChange={(e) => updateHotelRoomLine(line.id, { netPrice: e.target.value })}
                        placeholder="0"
                      />
                    </div>
                    <div>
                      <label className="text-[11px] text-stone-500 block mb-1">Sold (per room/night)</label>
                      <input
                        type="number"
                        className="w-28 border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 price-input"
                        value={line.soldPrice}
                        onChange={(e) => updateHotelRoomLine(line.id, { soldPrice: e.target.value })}
                        placeholder="0"
                      />
                      <div className="flex items-center justify-between gap-2 mt-3">
                        <div className="text-xs text-emerald-700 font-semibold">
                          {roomLineNights(line, hotelForm)} night{roomLineNights(line, hotelForm) === 1 ? "" : "s"} · {fmt(hotelLineSoldTotal(line, roomLineNights(line, hotelForm)) - hotelLineNetTotal(line, roomLineNights(line, hotelForm)))} {line.currency}
                        </div>
                        <button
                          onClick={() => removeHotelRoomLine(line.id)}
                          disabled={hotelForm.roomLines.length <= 1}
                          className="text-red-500 hover:text-red-700 disabled:opacity-30"
                          title="Remove this room line"
                        >
                          <Trash2 size={16} />
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* Adult guest names — one field per bed the room type holds, placed
                      directly above the Children section. Only the first guest is
                      mandatory; the rest are optional. */}
                  <div className="space-y-2">
                    {(line.guests || []).map((g, i) => (
                      <div key={g.id} className="bg-white border border-stone-200 rounded-lg p-2">
                        <label className="text-[11px] text-stone-500 block mb-1">
                          Guest {i + 1} name
                          {i === 0 ? <span className="text-red-500"> *</span> : (
                            <span className="text-stone-400"> (optional)</span>
                          )}
                        </label>
                        <input
                          className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                          value={g.name}
                          onChange={(e) => updateRoomGuest(line.id, i, e.target.value)}
                          placeholder={i === 0 ? "Guest 1 name (required)" : `Guest ${i + 1} name`}
                        />
                      </div>
                    ))}
                  </div>

                  {/* Children in this room — name + age in years (0–11). */}
                  <div>
                    <div className="flex items-center justify-between mb-1">
                      <label className="text-[11px] text-stone-500 block">Children</label>
                      <button
                        type="button"
                        onClick={() => addRoomChild(line.id)}
                        className="text-[11px] font-semibold text-teal-800 border border-teal-700 border-dashed rounded-lg px-2 py-1 hover:bg-teal-50"
                      >
                        + Add child
                      </button>
                    </div>
                    {(line.children || []).length > 0 && (
                      <div className="space-y-2">
                        {line.children.map((c, i) => (
                          <div key={c.id} className="grid grid-cols-1 sm:grid-cols-8 gap-3 items-end bg-white border border-stone-200 rounded-lg p-3">
                            <div className="sm:col-span-6">
                              <label className="text-[11px] text-stone-500 block mb-1">Child {i + 1} name</label>
                              <input
                                className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                                value={c.name}
                                onChange={(e) => updateRoomChild(line.id, c.id, { name: e.target.value })}
                                placeholder="Child name"
                              />
                            </div>
                            <div>
                              <label className="text-[11px] text-stone-500 block mb-1">Age (0–11)</label>
                              <input
                                type="text"
                                inputMode="numeric"
                                pattern="[0-9]*"
                                className="w-full border border-stone-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                                value={c.age}
                                onChange={(e) => updateRoomChild(line.id, c.id, { age: sanitizeAgeInput(e.target.value) })}
                                placeholder="e.g. 4"
                              />
                            </div>
                            <div className="flex justify-end">
                              <button
                                onClick={() => removeRoomChild(line.id, c.id)}
                                className="text-red-500 hover:text-red-700"
                                title="Remove this child"
                              >
                                <Trash2 size={14} />
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              ))}
              <button
                onClick={addHotelRoomLine}
                className="text-xs font-semibold text-teal-800 border border-teal-700 border-dashed rounded-lg px-3 py-1.5 hover:bg-teal-50"
              >
                + Add another room
              </button>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mt-4">
              <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                <p className="text-[11px] text-stone-500">Net total (EGP)</p>
                <p className="text-sm font-bold text-stone-800">{fmt(hotelNetTotal(hotelForm))}</p>
              </div>
              <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                <p className="text-[11px] text-stone-500">Sold total (EGP)</p>
                <p className="text-sm font-bold text-stone-800">{fmt(hotelSoldTotal(hotelForm))}</p>
              </div>
              <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                <p className="text-[11px] text-stone-500">Profit (auto, EGP)</p>
                <p className="text-sm font-bold text-emerald-700">{fmt(hotelProfitTotal(hotelForm))}</p>
              </div>
            </div>
            <div className="flex items-center gap-3 mt-4">
              <button
                onClick={handleSaveHotel}
                className="bg-gradient-to-b from-teal-700 to-teal-900 text-white text-sm font-semibold rounded-xl px-5 py-2.5 hover:brightness-110"
              >
                {hotelEditingId ? "Save changes" : "Add booking"}
              </button>
              {hotelEditingId && (
                <button
                  onClick={resetHotelForm}
                  className="text-sm font-semibold text-stone-500 rounded-xl px-4 py-2.5 hover:bg-stone-50"
                >
                  Cancel
                </button>
              )}
            </div>
          </div>
        )}

        {/* Search and filters — same unified card style as Flights, adapted to the
            fields hotel bookings actually have (no month/year select stub — those
            come from each booking's own booking date). */}
        <div className="bg-white border border-stone-200 rounded-2xl p-3 sm:p-4 mb-3">
          <div className="flex items-stretch gap-2">
            <div className="relative flex-1">
              <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
              <input
                className="w-full border border-stone-300 rounded-xl pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                placeholder="Search by employee, customer, hotel, or supplier"
                value={hotelQuery}
                onChange={(e) => setHotelQuery(e.target.value)}
              />
            </div>
            <button
              type="button"
              onClick={() => setHotelFiltersOpen(!hotelFiltersOpen)}
              className={`shrink-0 flex items-center gap-1.5 border rounded-xl px-3 py-2 text-sm font-medium transition-colors ${
                hotelFiltersOpen ? "border-teal-700 text-teal-800 bg-teal-50" : "border-stone-300 text-stone-600 hover:bg-stone-50 bg-white"
              }`}
            >
              <SlidersHorizontal size={16} />
              <span className="hidden sm:inline">Filters</span>
              {activeHotelFilterCount > 0 && (
                <span className="bg-teal-700 text-white text-[11px] font-bold rounded-full min-w-[20px] h-5 px-1 flex items-center justify-center">
                  {activeHotelFilterCount}
                </span>
              )}
              <ChevronDown size={14} className={`transition-transform ${hotelFiltersOpen ? "rotate-180" : ""}`} />
            </button>
          </div>

          {hotelFiltersOpen && (
            <div className="flex flex-wrap gap-3 mt-3 pt-3 border-t border-stone-100">
              <div>
                <label className="text-xs text-stone-500 block mb-1">Year</label>
                <MultiSelectDropdown
                  label="years"
                  icon={Calendar}
                  options={hotelYearsAvailable}
                  selected={hotelSelectedYear}
                  onChange={setHotelSelectedYear}
                  placeholder="All years"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Month</label>
                <MultiSelectDropdown
                  label="months"
                  icon={Calendar}
                  options={hotelMonthsAvailable.map((key) => ({ value: key, label: monthLabel(key) }))}
                  selected={hotelSelectedMonth}
                  onChange={setHotelSelectedMonth}
                  placeholder="All months"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Employee</label>
                <MultiSelectDropdown
                  label="employees"
                  icon={User}
                  options={hotelEmployeesAvailable}
                  selected={hotelSelectedEmployee}
                  onChange={setHotelSelectedEmployee}
                  placeholder="All employees"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Supplier</label>
                <MultiSelectDropdown
                  label="suppliers"
                  icon={Building2}
                  options={hotelSuppliersAvailable}
                  selected={hotelSelectedSupplier}
                  onChange={setHotelSelectedSupplier}
                  placeholder="All suppliers"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Hotel</label>
                <MultiSelectDropdown
                  label="hotels"
                  icon={Building2}
                  options={hotelNamesAvailable}
                  selected={hotelSelectedHotelName}
                  onChange={setHotelSelectedHotelName}
                  placeholder="All hotels"
                />
              </div>
            </div>
          )}

          {hasActiveHotelFilter && (
            <div className="flex flex-wrap items-center gap-x-4 gap-y-1.5 mt-3 pt-3 border-t border-stone-100 text-xs">
              <span className="text-stone-400 font-medium">Applied:</span>
              {hotelSelectedYear.map((y) => (
                <span key={`year-${y}`} className="inline-flex items-center gap-1 text-stone-600">
                  Year: <span className="font-semibold text-stone-800">{y}</span>
                  <button onClick={() => setHotelSelectedYear(hotelSelectedYear.filter((v) => v !== y))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              ))}
              {hotelSelectedMonth.map((m) => (
                <span key={`month-${m}`} className="inline-flex items-center gap-1 text-stone-600">
                  Month: <span className="font-semibold text-stone-800">{monthLabel(m)}</span>
                  <button onClick={() => setHotelSelectedMonth(hotelSelectedMonth.filter((v) => v !== m))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              ))}
              {hotelSelectedEmployee.map((e) => (
                <span key={`employee-${e}`} className="inline-flex items-center gap-1 text-stone-600">
                  Employee: <span className="font-semibold text-stone-800">{e}</span>
                  <button onClick={() => setHotelSelectedEmployee(hotelSelectedEmployee.filter((v) => v !== e))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              ))}
              {hotelSelectedSupplier.map((s) => (
                <span key={`supplier-${s}`} className="inline-flex items-center gap-1 text-stone-600">
                  Supplier: <span className="font-semibold text-stone-800">{s}</span>
                  <button onClick={() => setHotelSelectedSupplier(hotelSelectedSupplier.filter((v) => v !== s))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              ))}
              {hotelSelectedHotelName.map((h) => (
                <span key={`hotel-${h}`} className="inline-flex items-center gap-1 text-stone-600">
                  Hotel: <span className="font-semibold text-stone-800">{h}</span>
                  <button onClick={() => setHotelSelectedHotelName(hotelSelectedHotelName.filter((v) => v !== h))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              ))}
              {hotelQuery.trim() && (
                <span className="inline-flex items-center gap-1 text-stone-600">
                  Search: <span className="font-semibold text-stone-800">"{hotelQuery.trim()}"</span>
                  <button onClick={() => setHotelQuery("")} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              )}
              <button onClick={clearAllHotelFilters} className="text-red-600 hover:text-red-700 font-semibold ml-auto">
                Clear all
              </button>
            </div>
          )}
        </div>

        <div className="bg-white border border-stone-200 rounded-2xl overflow-x-auto">
          <table className="w-full text-xs border-collapse">
            <thead>
              <tr className="bg-stone-50 border-b border-stone-200 text-stone-500">
                <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Company</th>
                <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Hotel</th>
                <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Supplier</th>
                <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Rooms</th>
                <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap"># rooms</th>
                <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Booking date</th>
                <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Dates</th>
                <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Net total (EGP)</th>
                <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Sold total (EGP)</th>
                <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Profit (EGP)</th>
              </tr>
            </thead>
            <tbody>
              {filteredHotelBookings.length === 0 && (
                <tr>
                  <td colSpan={10} className="text-center text-stone-400 px-2.5 py-6">
                    {visibleHotelBookings.length === 0 ? "No hotel bookings yet." : "No hotel bookings match the current search/filters."}
                  </td>
                </tr>
              )}
              {filteredHotelBookings.map((h) => (
                <tr
                  key={h.id}
                  className="border-b border-stone-100 hover:bg-stone-50 cursor-pointer"
                  onClick={() => setViewingHotelBooking(h)}
                >
                  <td className="px-2.5 py-1 text-stone-700 whitespace-nowrap">
                    {h.customer && h.customer.trim() ? (
                      h.customer
                    ) : (
                      <span className="text-stone-400 italic">Individual</span>
                    )}
                  </td>
                  <td className="px-2.5 py-1 text-stone-700 whitespace-nowrap">{h.hotel}</td>
                  <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">{h.supplier}</td>
                  <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">{hotelLinesSummary(h)}</td>
                  <td className="px-2.5 py-1 text-stone-600 text-right whitespace-nowrap">{hotelRoomCount(h)}</td>
                  <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">
                    {h.bookingDate ? formatDisplayDate(h.bookingDate) : "-"}
                  </td>
                  <td className="px-2.5 py-1 text-stone-600 whitespace-nowrap">
                    {hotelDateRange(h).start && hotelDateRange(h).end
                      ? `${formatDisplayDate(hotelDateRange(h).start)} → ${formatDisplayDate(hotelDateRange(h).end)}`
                      : "-"}
                  </td>
                  <td className="px-2.5 py-1 text-stone-600 text-right whitespace-nowrap">{fmt(hotelNetTotal(h))}</td>
                  <td className="px-2.5 py-1 text-stone-600 text-right whitespace-nowrap">{fmt(hotelSoldTotal(h))}</td>
                  <td className="px-2.5 py-1 font-semibold text-emerald-700 text-right whitespace-nowrap">
                    {fmt(hotelProfitTotal(h))}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {viewingHotelBooking && (
          <div
            className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4"
            onClick={() => setViewingHotelBooking(null)}
          >
            <div
              className="bg-white rounded-2xl max-w-3xl w-full max-h-[85vh] overflow-y-auto p-6"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="flex items-start justify-between mb-4">
                <div>
                  <h3 className="text-lg font-bold text-stone-800">{viewingHotelBooking.hotel}</h3>
                  <p className="text-sm text-stone-500">
                    {viewingHotelBooking.customer && viewingHotelBooking.customer.trim() ? (
                      <>Company: {viewingHotelBooking.customer} <span className="text-teal-700 font-semibold">(Corporate)</span></>
                    ) : (
                      <span className="italic">Individual booking</span>
                    )}
                  </p>
                </div>
                <div className="flex items-center gap-1 shrink-0">
                  <button
                    onClick={() => handlePrintHotel(viewingHotelBooking)}
                    className="text-stone-400 hover:text-teal-800 p-1.5"
                    title="Print"
                  >
                    <Printer size={18} />
                  </button>
                  <button
                    onClick={() => setCopyPickerSource({ type: "hotels", record: viewingHotelBooking })}
                    className="text-stone-400 hover:text-amber-600 p-1.5"
                    title="Copy to a file"
                  >
                    <FileText size={18} />
                  </button>
                  {hotelsPerm.canEdit && (
                    <button
                      onClick={() => { handleEditHotelClick(viewingHotelBooking); setViewingHotelBooking(null); }}
                      className="text-stone-400 hover:text-teal-800 p-1.5"
                      title="Edit"
                    >
                      <Pencil size={18} />
                    </button>
                  )}
                  {hotelsPerm.canDelete && (
                    <button
                      onClick={() => {
                        const id = viewingHotelBooking.id;
                        handleDeleteHotel(id, () => setViewingHotelBooking(null));
                      }}
                      className="text-stone-400 hover:text-red-600 p-1.5"
                      title="Delete"
                    >
                      <Trash2 size={18} />
                    </button>
                  )}
                  <button
                    onClick={() => setViewingHotelBooking(null)}
                    className="text-stone-400 hover:text-stone-700 p-1.5"
                    title="Close"
                  >
                    <X size={20} />
                  </button>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4 text-sm">
                <div><span className="text-stone-500">Supplier: </span>{viewingHotelBooking.supplier || "-"}</div>
                <div><span className="text-stone-500">Booking date: </span>{viewingHotelBooking.bookingDate ? formatDisplayDate(viewingHotelBooking.bookingDate) : "-"}</div>
                <div><span className="text-stone-500">Employee: </span>{viewingHotelBooking.employee || "-"}</div>
                <div><span className="text-stone-500">Notes: </span>{viewingHotelBooking.notes || "-"}</div>
              </div>

              <div className="space-y-3">
                {(viewingHotelBooking.roomLines || []).map((l, idx) => {
                  const type = ROOM_TYPES.find((r) => r.value === l.roomType)?.label || l.roomType;
                  const meal = MEAL_PLANS.find((m) => m.value === l.mealPlan)?.label || l.mealPlan;
                  const nights = roomLineNights(l, viewingHotelBooking);
                  return (
                    <div key={l.id || idx} className="border border-stone-200 rounded-xl p-3">
                      <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
                        <span className="font-semibold text-stone-700 text-sm">
                          {l.count}× {type} — {meal}
                        </span>
                        <span className="text-xs text-stone-500">
                          {l.checkIn ? formatDisplayDate(l.checkIn) : "-"} → {l.checkOut ? formatDisplayDate(l.checkOut) : "-"} ({nights} night{nights === 1 ? "" : "s"})
                        </span>
                      </div>
                      <div className="text-xs text-stone-600 mb-2">
                        Net: {fmt(hotelLineNetTotal(l, nights))} {l.currency} · Sold:{" "}
                        {fmt(hotelLineSoldTotal(l, nights))} {l.currency}
                      </div>
                      {Array.isArray(l.guests) && l.guests.some((g) => g.name) && (
                        <div className="text-xs text-stone-700 mb-1">
                          <span className="text-stone-500">Guests: </span>
                          {l.guests.map((g) => g.name || "-").join(", ")}
                        </div>
                      )}
                      {Array.isArray(l.children) && l.children.length > 0 && (
                        <div className="text-xs text-stone-700">
                          <span className="text-stone-500">Children: </span>
                          {l.children
                            .map((c) => `${c.name || "-"} (${c.age !== "" && c.age != null ? c.age : "-"}y)`)
                            .join(", ")}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mt-4">
                <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                  <p className="text-[11px] text-stone-500">Net total (EGP)</p>
                  <p className="text-sm font-bold text-stone-800">{fmt(hotelNetTotal(viewingHotelBooking))}</p>
                </div>
                <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                  <p className="text-[11px] text-stone-500">Sold total (EGP)</p>
                  <p className="text-sm font-bold text-stone-800">{fmt(hotelSoldTotal(viewingHotelBooking))}</p>
                </div>
                <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                  <p className="text-[11px] text-stone-500">Profit (EGP)</p>
                  <p className="text-sm font-bold text-emerald-700">{fmt(hotelProfitTotal(viewingHotelBooking))}</p>
                </div>
              </div>
            </div>
          </div>
        )}
        </>
        )}

        {activeSection === "visa" && (
        <>
        {/* Summary cards, same style as the Flights section */}
        <div className="grid grid-cols-3 gap-1.5 sm:gap-3 mb-6">
          <div className="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0">
            <div className="bg-stone-100 rounded-xl p-1.5 sm:p-2 text-stone-600 shrink-0"><PassportIcon size={18} className="sm:hidden" /><PassportIcon size={20} className="hidden sm:block" /></div>
            <div className="min-w-0">
              <p className="text-xs text-stone-500">Applicants</p>
              <p className="text-sm sm:text-lg font-bold truncate">{visaTotals.count}</p>
            </div>
          </div>
          <div className="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0">
            <div className="bg-teal-50 rounded-xl p-1.5 sm:p-2 text-teal-900 shrink-0"><Wallet size={18} className="sm:hidden" /><Wallet size={20} className="hidden sm:block" /></div>
            <div className="min-w-0">
              <p className="text-xs text-stone-500">Total sales (EGP)</p>
              <p className="text-sm sm:text-lg font-bold truncate">{fmt(visaTotals.sold)}</p>
            </div>
          </div>
          <div className="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0">
            <div className="bg-emerald-50 rounded-xl p-1.5 sm:p-2 text-emerald-700 shrink-0"><TrendingUp size={18} className="sm:hidden" /><TrendingUp size={20} className="hidden sm:block" /></div>
            <div className="min-w-0">
              <p className="text-xs text-stone-500">Total profit (EGP)</p>
              <p className="text-sm sm:text-lg font-bold text-emerald-700 truncate">{fmt(visaTotals.profit)}</p>
            </div>
          </div>
        </div>

        {/* Button to register new supplier names for the Visa page's own supplier list —
            kept separate from the Hotels/Flights supplier lists. */}
        <div className="flex flex-wrap items-center gap-2 mb-4">
          <button
            onClick={() => setShowAddVisaSupplierPanel(!showAddVisaSupplierPanel)}
            className="text-xs font-semibold text-teal-800 border border-teal-700 rounded-xl px-3 py-2 hover:bg-teal-50 flex items-center gap-1.5"
          >
            <Plus size={14} /> Add supplier
          </button>
        </div>

        {showAddVisaSupplierPanel && (
          <div className="bg-white border border-stone-200 rounded-2xl p-4 mb-4">
            <h3 className="text-sm font-bold text-stone-700 mb-3">Visa suppliers</h3>
            <div className="flex gap-2 mb-3">
              <input
                className="w-full max-w-xs border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={newVisaSupplierDraft}
                onChange={(e) => setNewVisaSupplierDraft(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleAddVisaSupplierName()}
                placeholder="Supplier name"
              />
              <button
                onClick={handleAddVisaSupplierName}
                className="bg-gradient-to-b from-teal-700 to-teal-900 text-white text-sm font-semibold rounded-xl px-4 py-2 hover:brightness-110"
              >
                Add
              </button>
            </div>
            {(suggestions.visaSuppliers || []).length === 0 ? (
              <p className="text-xs text-stone-400">No suppliers saved yet</p>
            ) : (
              <div className="flex flex-wrap gap-2">
                {(suggestions.visaSuppliers || []).map((s) => (
                  <span
                    key={s}
                    className="inline-flex items-center gap-1.5 bg-stone-50 border border-stone-200 rounded-lg px-2.5 py-1 text-xs text-stone-700"
                  >
                    {s}
                    <button onClick={() => handleDeleteVisaSupplierName(s)} className="text-red-500 hover:text-red-700">
                      <Trash2 size={12} />
                    </button>
                  </span>
                ))}
              </div>
            )}
          </div>
        )}

        {visaError && (
          <div className="text-sm rounded-xl px-3 py-2 mb-4 bg-red-50 text-red-700">{visaError}</div>
        )}

        {canAddTickets && (
          <div className="bg-white border border-stone-200 rounded-2xl p-5 mb-6">
            <h3 className="text-sm font-bold text-stone-700 mb-4">
              {visaEditingId ? "Edit visa booking" : "New visa booking"}
            </h3>
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-4">
              <div>
                <label className="text-xs text-stone-500 block mb-1">Number of customers</label>
                <input
                  type="number"
                  min={1}
                  max={50}
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={visaForm.customersCount}
                  onChange={(e) => handleVisaCustomersCountChange(e.target.value)}
                  onBlur={(e) => {
                    if (e.target.value === "" || parseInt(e.target.value, 10) < 1) {
                      handleVisaCustomersCountChange(1);
                    }
                  }}
                  placeholder="1"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Visa</label>
                <input
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={visaForm.visaType}
                  onChange={(e) => setVisaForm({ ...visaForm, visaType: e.target.value })}
                  placeholder="e.g. Schengen, UK, Dubai"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Booking date</label>
                <input
                  type="date"
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={visaForm.bookingDate}
                  onChange={(e) => setVisaForm({ ...visaForm, bookingDate: e.target.value })}
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Supplier</label>
                {visaSupplierOther ? (
                  <div className="flex gap-2">
                    <input
                      className={`w-full border rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 ${visaForm.supplier.trim() ? "border-blue-400 text-blue-700 font-medium bg-blue-50" : "border-stone-300"}`}
                      value={visaForm.supplier}
                      onChange={(e) => setVisaForm({ ...visaForm, supplier: e.target.value })}
                      placeholder="Enter supplier name"
                      autoFocus
                    />
                    <button
                      type="button"
                      onClick={() => { setVisaSupplierOther(false); setVisaForm({ ...visaForm, supplier: "" }); }}
                      className="shrink-0 text-xs text-stone-500 hover:text-teal-800 border border-stone-300 rounded-xl px-2"
                    >
                      List
                    </button>
                  </div>
                ) : (
                  <select
                    className={`w-full border rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 ${visaForm.supplier ? "border-blue-400 text-blue-700 font-medium bg-blue-50" : "border-stone-300 bg-white"}`}
                    value={visaForm.supplier}
                    onChange={(e) => {
                      if (e.target.value === "__other__") {
                        setVisaSupplierOther(true);
                        setVisaForm({ ...visaForm, supplier: "" });
                      } else {
                        setVisaForm({ ...visaForm, supplier: e.target.value });
                      }
                    }}
                  >
                    <option value="">Select supplier</option>
                    {(suggestions.visaSuppliers || []).map((s) => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                    <option value="__other__">Other</option>
                  </select>
                )}
              </div>
            </div>


            {/* Dynamic customer name cells, one row per customer */}
            <div className="mb-4">
              <label className="text-xs text-stone-500 block mb-2">
                Customers ({visaForm.customers.length})
              </label>
              <div className="space-y-2">
                {visaForm.customers.map((c, i) => (
                  <input
                    key={i}
                    className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                    value={c.name}
                    onChange={(e) => handleVisaCustomerNameChange(i, e.target.value)}
                    placeholder={i === 0 ? `Customer ${i + 1} name (required)` : `Customer ${i + 1} name`}
                  />
                ))}
              </div>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-3 gap-4 mb-4">
              <div>
                <label className="text-xs text-stone-500 block mb-1">Currency</label>
                <select
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white"
                  value={visaForm.currency}
                  onChange={(e) => setVisaForm({ ...visaForm, currency: e.target.value })}
                >
                  {HOTEL_CURRENCIES.map((c) => (
                    <option key={c.value} value={c.value}>{c.label}</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Price net (per person)</label>
                <input
                  type="number"
                  className="w-28 border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 price-input"
                  value={visaForm.netPrice}
                  onChange={(e) => setVisaForm({ ...visaForm, netPrice: e.target.value })}
                  placeholder="0"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Sold (per person)</label>
                <input
                  type="number"
                  className="w-28 border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 price-input"
                  value={visaForm.soldPrice}
                  onChange={(e) => setVisaForm({ ...visaForm, soldPrice: e.target.value })}
                  placeholder="0"
                />
              </div>
            </div>

            {/* Live total preview: per-person prices above multiplied by the number of
                customers entered, same style as the Hotels form's totals box. */}
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mt-1 mb-4">
              <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                <p className="text-[11px] text-stone-500">Net total (× {visaForm.customers.length || 1})</p>
                <p className="text-sm font-bold text-stone-800">{fmt(visaNetTotal(visaForm))} {visaForm.currency}</p>
              </div>
              <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                <p className="text-[11px] text-stone-500">Sold total (× {visaForm.customers.length || 1})</p>
                <p className="text-sm font-bold text-stone-800">{fmt(visaSoldTotal(visaForm))} {visaForm.currency}</p>
              </div>
              <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                <p className="text-[11px] text-stone-500">Profit (auto)</p>
                <p className="text-sm font-bold text-emerald-700">{fmt(visaProfitTotal(visaForm))} {visaForm.currency}</p>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={handleSaveVisa}
                className="bg-gradient-to-b from-teal-700 to-teal-900 text-white text-sm font-semibold rounded-xl px-4 py-2 hover:brightness-110"
              >
                {visaEditingId ? "Save changes" : "Add visa booking"}
              </button>
              {visaEditingId && (
                <button
                  onClick={resetVisaForm}
                  className="text-sm text-stone-500 hover:text-stone-700 border border-stone-300 rounded-xl px-4 py-2"
                >
                  Cancel
                </button>
              )}
            </div>
          </div>
        )}

        {/* Search and filters — same unified card style as Flights/Hotels, adapted to
            the fields visa bookings actually have (no Employee filter — visa bookings
            don't track which employee created them). */}
        <div className="bg-white border border-stone-200 rounded-2xl p-3 sm:p-4 mb-3">
          <div className="flex items-stretch gap-2">
            <div className="relative flex-1">
              <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
              <input
                className="w-full border border-stone-300 rounded-xl pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                placeholder="Search by customer name, visa type, or supplier"
                value={visaQuery}
                onChange={(e) => setVisaQuery(e.target.value)}
              />
            </div>
            <button
              type="button"
              onClick={() => setVisaFiltersOpen(!visaFiltersOpen)}
              className={`shrink-0 flex items-center gap-1.5 border rounded-xl px-3 py-2 text-sm font-medium transition-colors ${
                visaFiltersOpen ? "border-teal-700 text-teal-800 bg-teal-50" : "border-stone-300 text-stone-600 hover:bg-stone-50 bg-white"
              }`}
            >
              <SlidersHorizontal size={16} />
              <span className="hidden sm:inline">Filters</span>
              {activeVisaFilterCount > 0 && (
                <span className="bg-teal-700 text-white text-[11px] font-bold rounded-full min-w-[20px] h-5 px-1 flex items-center justify-center">
                  {activeVisaFilterCount}
                </span>
              )}
              <ChevronDown size={14} className={`transition-transform ${visaFiltersOpen ? "rotate-180" : ""}`} />
            </button>
          </div>

          {visaFiltersOpen && (
            <div className="flex flex-wrap gap-3 mt-3 pt-3 border-t border-stone-100">
              <div>
                <label className="text-xs text-stone-500 block mb-1">Year</label>
                <MultiSelectDropdown
                  label="years"
                  icon={Calendar}
                  options={visaYearsAvailable}
                  selected={visaSelectedYear}
                  onChange={setVisaSelectedYear}
                  placeholder="All years"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Month</label>
                <MultiSelectDropdown
                  label="months"
                  icon={Calendar}
                  options={visaMonthsAvailable.map((key) => ({ value: key, label: monthLabel(key) }))}
                  selected={visaSelectedMonth}
                  onChange={setVisaSelectedMonth}
                  placeholder="All months"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Supplier</label>
                <MultiSelectDropdown
                  label="suppliers"
                  icon={Building2}
                  options={visaSuppliersAvailable}
                  selected={visaSelectedSupplier}
                  onChange={setVisaSelectedSupplier}
                  placeholder="All suppliers"
                />
              </div>
            </div>
          )}

          {hasActiveVisaFilter && (
            <div className="flex flex-wrap items-center gap-x-4 gap-y-1.5 mt-3 pt-3 border-t border-stone-100 text-xs">
              <span className="text-stone-400 font-medium">Applied:</span>
              {visaSelectedYear.map((y) => (
                <span key={`year-${y}`} className="inline-flex items-center gap-1 text-stone-600">
                  Year: <span className="font-semibold text-stone-800">{y}</span>
                  <button onClick={() => setVisaSelectedYear(visaSelectedYear.filter((v) => v !== y))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              ))}
              {visaSelectedMonth.map((m) => (
                <span key={`month-${m}`} className="inline-flex items-center gap-1 text-stone-600">
                  Month: <span className="font-semibold text-stone-800">{monthLabel(m)}</span>
                  <button onClick={() => setVisaSelectedMonth(visaSelectedMonth.filter((v) => v !== m))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              ))}
              {visaSelectedSupplier.map((s) => (
                <span key={`supplier-${s}`} className="inline-flex items-center gap-1 text-stone-600">
                  Supplier: <span className="font-semibold text-stone-800">{s}</span>
                  <button onClick={() => setVisaSelectedSupplier(visaSelectedSupplier.filter((v) => v !== s))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              ))}
              {visaQuery.trim() && (
                <span className="inline-flex items-center gap-1 text-stone-600">
                  Search: <span className="font-semibold text-stone-800">"{visaQuery.trim()}"</span>
                  <button onClick={() => setVisaQuery("")} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              )}
              <button onClick={clearAllVisaFilters} className="text-red-600 hover:text-red-700 font-semibold ml-auto">
                Clear all
              </button>
            </div>
          )}
        </div>

        {/* Visa bookings list */}
        {filteredVisaBookings.length === 0 ? (
          <div className="bg-white border border-stone-200 rounded-2xl p-12 text-center text-stone-400">
            <PassportIcon size={40} className="mx-auto mb-3 text-stone-300" />
            <p className="text-sm">{visibleVisaBookings.length === 0 ? "No visa bookings yet." : "No visa bookings match the current search/filters."}</p>
          </div>
        ) : (
          <div className="bg-white border border-stone-200 rounded-2xl overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead className="bg-stone-50 text-stone-500 text-xs">
                  <tr>
                    <th className="text-left px-4 py-3 font-semibold"># Customers</th>
                    <th className="text-left px-4 py-3 font-semibold">Names</th>
                    <th className="text-left px-4 py-3 font-semibold">Visa</th>
                    <th className="text-left px-4 py-3 font-semibold">Booking date</th>
                    <th className="text-left px-4 py-3 font-semibold">Supplier</th>
                    <th className="text-right px-4 py-3 font-semibold">Net</th>
                    <th className="text-right px-4 py-3 font-semibold">Sold</th>
                    <th className="text-right px-4 py-3 font-semibold">Profit</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-stone-100">
                  {filteredVisaBookings.map((v) => {
                    const net = visaNetTotal(v);
                    const sold = visaSoldTotal(v);
                    const profit = sold - net;
                    return (
                      <tr
                        key={v.id}
                        className="hover:bg-stone-50 cursor-pointer"
                        onClick={() => setViewingVisaBooking(v)}
                      >
                        <td className="px-4 py-3 text-stone-700">{(v.customers || []).length}</td>
                        <td className="px-4 py-3 text-stone-700">
                          {(v.customers || []).map((c) => c.name || "-").join(", ")}
                        </td>
                        <td className="px-4 py-3 text-stone-700">{v.visaType}</td>
                        <td className="px-4 py-3 text-stone-700 whitespace-nowrap">
                          {v.bookingDate ? formatDisplayDate(v.bookingDate) : "-"}
                        </td>
                        <td className="px-4 py-3 text-stone-700">{v.supplier}</td>
                        <td className="px-4 py-3 text-right text-stone-700">{fmt(net)} {v.currency}</td>
                        <td className="px-4 py-3 text-right text-stone-700">{fmt(sold)} {v.currency}</td>
                        <td className="px-4 py-3 text-right font-semibold text-emerald-700">{fmt(profit)} {v.currency}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {viewingVisaBooking && (
          <div
            className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4"
            onClick={() => setViewingVisaBooking(null)}
          >
            <div
              className="bg-white rounded-2xl max-w-lg w-full max-h-[85vh] overflow-y-auto p-6"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="flex items-start justify-between mb-4">
                <div>
                  <h3 className="text-lg font-bold text-stone-800">{viewingVisaBooking.visaType || "Visa"}</h3>
                  <p className="text-sm text-stone-500">
                    {(viewingVisaBooking.customers || []).length} customer
                    {(viewingVisaBooking.customers || []).length === 1 ? "" : "s"}
                  </p>
                </div>
                <div className="flex items-center gap-1 shrink-0">
                  <button
                    onClick={() => handlePrintVisa(viewingVisaBooking)}
                    className="text-stone-400 hover:text-teal-800 p-1.5"
                    title="Print"
                  >
                    <Printer size={18} />
                  </button>
                  <button
                    onClick={() => setCopyPickerSource({ type: "visa", record: viewingVisaBooking })}
                    className="text-stone-400 hover:text-amber-600 p-1.5"
                    title="Copy to a file"
                  >
                    <FileText size={18} />
                  </button>
                  {visaPerm.canEdit && (
                    <button
                      onClick={() => { handleEditVisaClick(viewingVisaBooking); setViewingVisaBooking(null); }}
                      className="text-stone-400 hover:text-teal-800 p-1.5"
                      title="Edit"
                    >
                      <Pencil size={18} />
                    </button>
                  )}
                  {visaPerm.canDelete && (
                    <button
                      onClick={() => {
                        const id = viewingVisaBooking.id;
                        handleDeleteVisa(id, () => setViewingVisaBooking(null));
                      }}
                      className="text-stone-400 hover:text-red-600 p-1.5"
                      title="Delete"
                    >
                      <Trash2 size={18} />
                    </button>
                  )}
                  <button
                    onClick={() => setViewingVisaBooking(null)}
                    className="text-stone-400 hover:text-stone-700 p-1.5"
                    title="Close"
                  >
                    <X size={20} />
                  </button>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4 text-sm">
                <div><span className="text-stone-500">Supplier: </span>{viewingVisaBooking.supplier || "-"}</div>
                <div>
                  <span className="text-stone-500">Booking date: </span>
                  {viewingVisaBooking.bookingDate ? formatDisplayDate(viewingVisaBooking.bookingDate) : "-"}
                </div>
              </div>

              <div className="border border-stone-200 rounded-xl p-3 mb-4">
                <p className="text-xs text-stone-500 mb-1.5">Customers</p>
                <div className="text-sm text-stone-700 space-y-1">
                  {(viewingVisaBooking.customers || []).map((c, i) => (
                    <div key={i}>{i + 1}. {c.name || "-"}</div>
                  ))}
                </div>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                  <p className="text-[11px] text-stone-500">Net total</p>
                  <p className="text-sm font-bold text-stone-800">
                    {fmt(visaNetTotal(viewingVisaBooking))} {viewingVisaBooking.currency}
                  </p>
                </div>
                <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                  <p className="text-[11px] text-stone-500">Sold total</p>
                  <p className="text-sm font-bold text-stone-800">
                    {fmt(visaSoldTotal(viewingVisaBooking))} {viewingVisaBooking.currency}
                  </p>
                </div>
                <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                  <p className="text-[11px] text-stone-500">Profit</p>
                  <p className="text-sm font-bold text-emerald-700">
                    {fmt(visaProfitTotal(viewingVisaBooking))} {viewingVisaBooking.currency}
                  </p>
                </div>
              </div>
            </div>
          </div>
        )}
        </>
        )}

        {activeSection === "cars" && (
        <>
        {/* Summary cards, same style as the Flights section */}
        <div className="grid grid-cols-3 gap-1.5 sm:gap-3 mb-6">
          <div className="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0">
            <div className="bg-stone-100 rounded-xl p-1.5 sm:p-2 text-stone-600 shrink-0"><Car size={18} className="sm:hidden" /><Car size={20} className="hidden sm:block" /></div>
            <div className="min-w-0">
              <p className="text-xs text-stone-500">Bookings</p>
              <p className="text-sm sm:text-lg font-bold truncate">{carTotals.count}</p>
            </div>
          </div>
          <div className="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0">
            <div className="bg-teal-50 rounded-xl p-1.5 sm:p-2 text-teal-900 shrink-0"><Wallet size={18} className="sm:hidden" /><Wallet size={20} className="hidden sm:block" /></div>
            <div className="min-w-0">
              <p className="text-xs text-stone-500">Total sales (EGP)</p>
              <p className="text-sm sm:text-lg font-bold truncate">{fmt(carTotals.sold)}</p>
            </div>
          </div>
          <div className="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0">
            <div className="bg-emerald-50 rounded-xl p-1.5 sm:p-2 text-emerald-700 shrink-0"><TrendingUp size={18} className="sm:hidden" /><TrendingUp size={20} className="hidden sm:block" /></div>
            <div className="min-w-0">
              <p className="text-xs text-stone-500">Total profit (EGP)</p>
              <p className="text-sm sm:text-lg font-bold text-emerald-700 truncate">{fmt(carTotals.profit)}</p>
            </div>
          </div>
        </div>

        {/* Button to register new supplier names for the Transfers page's own supplier
            list — kept separate from the Hotels/Flights/Visa supplier lists. */}
        <div className="flex flex-wrap items-center gap-2 mb-4">
          <button
            onClick={() => setShowAddCarSupplierPanel(!showAddCarSupplierPanel)}
            className="text-xs font-semibold text-teal-800 border border-teal-700 rounded-xl px-3 py-2 hover:bg-teal-50 flex items-center gap-1.5"
          >
            <Plus size={14} /> Add supplier
          </button>
        </div>

        {showAddCarSupplierPanel && (
          <div className="bg-white border border-stone-200 rounded-2xl p-4 mb-4">
            <h3 className="text-sm font-bold text-stone-700 mb-3">Transfer suppliers</h3>
            <div className="flex gap-2 mb-3">
              <input
                className="w-full max-w-xs border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                value={newCarSupplierDraft}
                onChange={(e) => setNewCarSupplierDraft(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleAddCarSupplierName()}
                placeholder="Supplier name"
              />
              <button
                onClick={handleAddCarSupplierName}
                className="bg-gradient-to-b from-teal-700 to-teal-900 text-white text-sm font-semibold rounded-xl px-4 py-2 hover:brightness-110"
              >
                Add
              </button>
            </div>
            {(suggestions.carSuppliers || []).length === 0 ? (
              <p className="text-xs text-stone-400">No suppliers saved yet</p>
            ) : (
              <div className="flex flex-wrap gap-2">
                {(suggestions.carSuppliers || []).map((s) => (
                  <span
                    key={s}
                    className="inline-flex items-center gap-1.5 bg-stone-50 border border-stone-200 rounded-lg px-2.5 py-1 text-xs text-stone-700"
                  >
                    {s}
                    <button onClick={() => handleDeleteCarSupplierName(s)} className="text-red-500 hover:text-red-700">
                      <Trash2 size={12} />
                    </button>
                  </span>
                ))}
              </div>
            )}
          </div>
        )}

        {carError && (
          <div className="text-sm rounded-xl px-3 py-2 mb-4 bg-red-50 text-red-700">{carError}</div>
        )}

        {canAddTickets && (
          <div className="bg-white border border-stone-200 rounded-2xl p-5 mb-6">
            <h3 className="text-sm font-bold text-stone-700 mb-4">
              {carEditingId ? "Edit transfer booking" : "New transfer booking"}
            </h3>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
              <div>
                <label className="text-xs text-stone-500 block mb-1">Customer name</label>
                <input
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={carForm.customerName}
                  onChange={(e) => setCarForm({ ...carForm, customerName: e.target.value })}
                  placeholder="Customer name"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Phone number</label>
                <input
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={carForm.phone}
                  onChange={(e) => setCarForm({ ...carForm, phone: e.target.value })}
                  placeholder="Phone number"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Entry date (booking entered on)</label>
                <input
                  type="date"
                  max={todayDateStr()}
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={carForm.entryDate}
                  onChange={(e) => setCarForm({ ...carForm, entryDate: e.target.value })}
                />
              </div>
            </div>


            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
              <div>
                <label className="text-xs text-stone-500 block mb-1">Route — from</label>
                <input
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={carForm.routeFrom}
                  onChange={(e) => setCarForm({ ...carForm, routeFrom: e.target.value })}
                  placeholder="e.g. Cairo Airport"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Route — to</label>
                <input
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={carForm.routeTo}
                  onChange={(e) => setCarForm({ ...carForm, routeTo: e.target.value })}
                  placeholder="e.g. Downtown Hotel"
                />
              </div>
            </div>

            {/* Pickup date & time — placed right after the route so the run's "where" and "when" read together */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
              <div>
                <label className="text-xs text-stone-500 block mb-1">Date</label>
                <input
                  type="date"
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                  value={carForm.bookingDate}
                  onChange={(e) => setCarForm({ ...carForm, bookingDate: e.target.value })}
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Time</label>
                <TimeSelect
                  value={carForm.bookingTime}
                  onChange={(v) => setCarForm({ ...carForm, bookingTime: v })}
                />
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
              <div>
                <label className="text-xs text-stone-500 block mb-1">Car type</label>
                <select
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white"
                  value={carForm.carType}
                  onChange={(e) => setCarForm({ ...carForm, carType: e.target.value })}
                >
                  <option value="">Select car type</option>
                  {CAR_TYPES.map((t) => (
                    <option key={t} value={t}>{t}</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Supplier</label>
                {carSupplierOther ? (
                  <div className="flex gap-2">
                    <input
                      className={`w-full border rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 ${carForm.supplier.trim() ? "border-blue-400 text-blue-700 font-medium bg-blue-50" : "border-stone-300"}`}
                      value={carForm.supplier}
                      onChange={(e) => setCarForm({ ...carForm, supplier: e.target.value })}
                      placeholder="Enter supplier name"
                      autoFocus
                    />
                    <button
                      type="button"
                      onClick={() => { setCarSupplierOther(false); setCarForm({ ...carForm, supplier: "" }); }}
                      className="shrink-0 text-xs text-stone-500 hover:text-teal-800 border border-stone-300 rounded-xl px-2"
                    >
                      List
                    </button>
                  </div>
                ) : (
                  <select
                    className={`w-full border rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 ${carForm.supplier ? "border-blue-400 text-blue-700 font-medium bg-blue-50" : "border-stone-300 bg-white"}`}
                    value={carForm.supplier}
                    onChange={(e) => {
                      if (e.target.value === "__other__") {
                        setCarSupplierOther(true);
                        setCarForm({ ...carForm, supplier: "" });
                      } else {
                        setCarForm({ ...carForm, supplier: e.target.value });
                      }
                    }}
                  >
                    <option value="">Select supplier</option>
                    {(suggestions.carSuppliers || []).map((s) => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                    <option value="__other__">Other</option>
                  </select>
                )}
              </div>
            </div>

            {/* Waiting hours / round trip / starts at the airport — grouped together
                in one horizontal card */}
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-4 mb-4">
              <div>
                <label className="flex items-center gap-2 text-xs text-stone-500 mb-1 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={carForm.hasWaiting}
                    onChange={(e) => setCarForm({ ...carForm, hasWaiting: e.target.checked, waitingHours: e.target.checked ? carForm.waitingHours : "" })}
                    className="rounded border-stone-300"
                  />
                  Waiting hours
                </label>
                {carForm.hasWaiting && (
                  <input
                    type="number"
                    className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                    value={carForm.waitingHours}
                    onChange={(e) => setCarForm({ ...carForm, waitingHours: e.target.value })}
                    placeholder="Number of hours"
                  />
                )}
              </div>
              <div>
                <label className="flex items-center gap-2 text-xs text-stone-500 mb-1 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={carForm.isRoundTrip}
                    onChange={(e) =>
                      setCarForm({
                        ...carForm,
                        isRoundTrip: e.target.checked,
                        returnDate: e.target.checked ? carForm.returnDate : "",
                        returnTime: e.target.checked ? carForm.returnTime : "",
                      })
                    }
                    className="rounded border-stone-300"
                  />
                  Round trip (go &amp; return)
                </label>
                <p className="text-xs text-stone-400 mt-2">
                  {carForm.isRoundTrip ? "Round trip" : "One way"}
                </p>
              </div>
              <div>
                <label className="flex items-center gap-2 text-xs text-stone-500 mb-1 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={carForm.startsAtAirport}
                    onChange={(e) => setCarForm({ ...carForm, startsAtAirport: e.target.checked, flightNumber: e.target.checked ? carForm.flightNumber : "" })}
                    className="rounded border-stone-300"
                  />
                  Starts at the airport
                </label>
                {carForm.startsAtAirport && (
                  <input
                    className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                    value={carForm.flightNumber}
                    onChange={(e) => setCarForm({ ...carForm, flightNumber: e.target.value })}
                    placeholder="Flight number"
                  />
                )}
              </div>
            </div>

            {/* Return date & time — only relevant for round trips */}
            {carForm.isRoundTrip && (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                <div>
                  <label className="text-xs text-stone-500 block mb-1">Return date</label>
                  <input
                    type="date"
                    className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                    value={carForm.returnDate}
                    onChange={(e) => setCarForm({ ...carForm, returnDate: e.target.value })}
                  />
                </div>
                <div>
                  <label className="text-xs text-stone-500 block mb-1">Return time</label>
                  <TimeSelect
                    value={carForm.returnTime}
                    onChange={(v) => setCarForm({ ...carForm, returnTime: v })}
                  />
                </div>
              </div>
            )}

            {/* Currency, amount to collect from the customer, net/sold prices, and driver tip */}
            <div className="grid grid-cols-2 sm:grid-cols-5 gap-4 mb-4">
              <div>
                <label className="text-xs text-stone-500 block mb-1">Currency</label>
                <select
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white"
                  value={carForm.currency}
                  onChange={(e) => setCarForm({ ...carForm, currency: e.target.value })}
                >
                  {HOTEL_CURRENCIES.map((c) => (
                    <option key={c.value} value={c.value}>{c.label}</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Collection</label>
                <input
                  type="number"
                  className="w-28 border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 price-input"
                  value={carForm.collection}
                  onChange={(e) => setCarForm({ ...carForm, collection: e.target.value })}
                  placeholder="0"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Price net</label>
                <input
                  type="number"
                  className="w-28 border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 price-input"
                  value={carForm.netPrice}
                  onChange={(e) => setCarForm({ ...carForm, netPrice: e.target.value })}
                  placeholder="0"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Sold</label>
                <input
                  type="number"
                  className="w-28 border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 price-input"
                  value={carForm.soldPrice}
                  onChange={(e) => setCarForm({ ...carForm, soldPrice: e.target.value })}
                  placeholder="0"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Driver tip</label>
                <input
                  type="number"
                  className="w-28 border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 price-input"
                  value={carForm.driverTip}
                  onChange={(e) => setCarForm({ ...carForm, driverTip: e.target.value })}
                  placeholder="0"
                />
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={handleSaveCar}
                className="bg-gradient-to-b from-teal-700 to-teal-900 text-white text-sm font-semibold rounded-xl px-4 py-2 hover:brightness-110"
              >
                {carEditingId ? "Save changes" : "Add transfer booking"}
              </button>
              {carEditingId && (
                <button
                  onClick={resetCarForm}
                  className="text-sm text-stone-500 hover:text-stone-700 border border-stone-300 rounded-xl px-4 py-2"
                >
                  Cancel
                </button>
              )}
            </div>
          </div>
        )}

        {/* Search and filters — same unified card style as the other sections, adapted
            to the fields transfer bookings actually have (no Employee filter — these
            bookings don't track which employee created them). */}
        <div className="bg-white border border-stone-200 rounded-2xl p-3 sm:p-4 mb-3">
          <div className="flex items-stretch gap-2">
            <div className="relative flex-1">
              <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
              <input
                className="w-full border border-stone-300 rounded-xl pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                placeholder="Search by customer, route, car type, supplier, or flight number"
                value={carQuery}
                onChange={(e) => setCarQuery(e.target.value)}
              />
            </div>
            <button
              type="button"
              onClick={() => setCarFiltersOpen(!carFiltersOpen)}
              className={`shrink-0 flex items-center gap-1.5 border rounded-xl px-3 py-2 text-sm font-medium transition-colors ${
                carFiltersOpen ? "border-teal-700 text-teal-800 bg-teal-50" : "border-stone-300 text-stone-600 hover:bg-stone-50 bg-white"
              }`}
            >
              <SlidersHorizontal size={16} />
              <span className="hidden sm:inline">Filters</span>
              {activeCarFilterCount > 0 && (
                <span className="bg-teal-700 text-white text-[11px] font-bold rounded-full min-w-[20px] h-5 px-1 flex items-center justify-center">
                  {activeCarFilterCount}
                </span>
              )}
              <ChevronDown size={14} className={`transition-transform ${carFiltersOpen ? "rotate-180" : ""}`} />
            </button>
          </div>

          {carFiltersOpen && (
            <div className="flex flex-wrap gap-3 mt-3 pt-3 border-t border-stone-100">
              <div>
                <label className="text-xs text-stone-500 block mb-1">Year</label>
                <MultiSelectDropdown
                  label="years"
                  icon={Calendar}
                  options={carYearsAvailable}
                  selected={carSelectedYear}
                  onChange={setCarSelectedYear}
                  placeholder="All years"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Month</label>
                <MultiSelectDropdown
                  label="months"
                  icon={Calendar}
                  options={carMonthsAvailable.map((key) => ({ value: key, label: monthLabel(key) }))}
                  selected={carSelectedMonth}
                  onChange={setCarSelectedMonth}
                  placeholder="All months"
                />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">Supplier</label>
                <MultiSelectDropdown
                  label="suppliers"
                  icon={Building2}
                  options={carSuppliersAvailable}
                  selected={carSelectedSupplier}
                  onChange={setCarSelectedSupplier}
                  placeholder="All suppliers"
                />
              </div>
            </div>
          )}

          {hasActiveCarFilter && (
            <div className="flex flex-wrap items-center gap-x-4 gap-y-1.5 mt-3 pt-3 border-t border-stone-100 text-xs">
              <span className="text-stone-400 font-medium">Applied:</span>
              {carSelectedYear.map((y) => (
                <span key={`year-${y}`} className="inline-flex items-center gap-1 text-stone-600">
                  Year: <span className="font-semibold text-stone-800">{y}</span>
                  <button onClick={() => setCarSelectedYear(carSelectedYear.filter((v) => v !== y))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              ))}
              {carSelectedMonth.map((m) => (
                <span key={`month-${m}`} className="inline-flex items-center gap-1 text-stone-600">
                  Month: <span className="font-semibold text-stone-800">{monthLabel(m)}</span>
                  <button onClick={() => setCarSelectedMonth(carSelectedMonth.filter((v) => v !== m))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              ))}
              {carSelectedSupplier.map((s) => (
                <span key={`supplier-${s}`} className="inline-flex items-center gap-1 text-stone-600">
                  Supplier: <span className="font-semibold text-stone-800">{s}</span>
                  <button onClick={() => setCarSelectedSupplier(carSelectedSupplier.filter((v) => v !== s))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              ))}
              {carQuery.trim() && (
                <span className="inline-flex items-center gap-1 text-stone-600">
                  Search: <span className="font-semibold text-stone-800">"{carQuery.trim()}"</span>
                  <button onClick={() => setCarQuery("")} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                </span>
              )}
              <button onClick={clearAllCarFilters} className="text-red-600 hover:text-red-700 font-semibold ml-auto">
                Clear all
              </button>
            </div>
          )}
        </div>

        {/* Transfer bookings list */}
        {filteredCarBookings.length === 0 ? (
          <div className="bg-white border border-stone-200 rounded-2xl p-12 text-center text-stone-400">
            <Car size={40} className="mx-auto mb-3 text-stone-300" />
            <p className="text-sm">{visibleCarBookings.length === 0 ? "No transfer bookings yet." : "No transfer bookings match the current search/filters."}</p>
          </div>
        ) : (
          <div className="bg-white border border-stone-200 rounded-2xl overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead className="bg-stone-50 text-stone-500 text-xs">
                  <tr>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Entry date</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Customer</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Phone</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Route</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Car type</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Supplier</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Trip</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Waiting</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Flight #</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Date &amp; time</th>
                    <th className="text-left px-2.5 py-1.5 font-semibold whitespace-nowrap">Return</th>
                    <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Collection</th>
                    <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Driver tip</th>
                    <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Net</th>
                    <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Sold</th>
                    <th className="text-right px-2.5 py-1.5 font-semibold whitespace-nowrap">Profit</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-stone-100">
                  {filteredCarBookings.map((c) => {
                    const net = parseFloat(c.netPrice) || 0;
                    const sold = parseFloat(c.soldPrice) || 0;
                    const profit = sold - net;
                    return (
                      <tr
                        key={c.id}
                        className="leading-tight hover:bg-stone-50 cursor-pointer"
                        onClick={() => setViewingCarBooking(c)}
                      >
                        <td className="px-2.5 py-1 text-stone-700 whitespace-nowrap">
                          {c.entryDate ? formatDisplayDate(c.entryDate) : "-"}
                        </td>
                        <td className="px-2.5 py-1 text-stone-700 whitespace-nowrap">{c.customerName}</td>
                        <td className="px-2.5 py-1 text-stone-700 whitespace-nowrap">{c.phone || "-"}</td>
                        <td className="px-2.5 py-1 text-stone-700 whitespace-nowrap">{c.routeFrom} → {c.routeTo}</td>
                        <td className="px-2.5 py-1 text-stone-700 whitespace-nowrap">{c.carType}</td>
                        <td className="px-2.5 py-1 text-stone-700 whitespace-nowrap">{c.supplier}</td>
                        <td className="px-2.5 py-1 text-stone-700 whitespace-nowrap">{c.isRoundTrip ? "Round trip" : "One way"}</td>
                        <td className="px-2.5 py-1 text-stone-700 whitespace-nowrap">
                          {c.hasWaiting ? `${c.waitingHours || 0} h` : "-"}
                        </td>
                        <td className="px-2.5 py-1 text-stone-700 whitespace-nowrap">
                          {c.startsAtAirport ? (c.flightNumber || "-") : "-"}
                        </td>
                        <td className="px-2.5 py-1 text-stone-700 whitespace-nowrap">
                          {c.bookingDate ? formatDisplayDate(c.bookingDate) : "-"}
                          {c.bookingTime ? ` · ${c.bookingTime}` : ""}
                        </td>
                        <td className="px-2.5 py-1 text-stone-700 whitespace-nowrap">
                          {c.isRoundTrip
                            ? `${c.returnDate ? formatDisplayDate(c.returnDate) : "-"}${c.returnTime ? ` · ${c.returnTime}` : ""}`
                            : "-"}
                        </td>
                        <td className="px-2.5 py-1 text-right text-stone-700 whitespace-nowrap">
                          {c.collection ? `${fmt(parseFloat(c.collection) || 0)} ${c.currency}` : "-"}
                        </td>
                        <td className="px-2.5 py-1 text-right text-stone-700 whitespace-nowrap">
                          {c.driverTip ? `${fmt(parseFloat(c.driverTip) || 0)} ${c.currency}` : "-"}
                        </td>
                        <td className="px-2.5 py-1 text-right text-stone-700 whitespace-nowrap">{fmt(net)} {c.currency}</td>
                        <td className="px-2.5 py-1 text-right text-stone-700 whitespace-nowrap">{fmt(sold)} {c.currency}</td>
                        <td className="px-2.5 py-1 text-right font-semibold text-emerald-700 whitespace-nowrap">{fmt(profit)} {c.currency}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {viewingCarBooking && (
          <div
            className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4"
            onClick={() => setViewingCarBooking(null)}
          >
            <div
              className="bg-white rounded-2xl max-w-lg w-full max-h-[85vh] overflow-y-auto p-6"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="flex items-start justify-between mb-4">
                <div>
                  <h3 className="text-lg font-bold text-stone-800">{viewingCarBooking.customerName || "Transfer"}</h3>
                  <p className="text-sm text-stone-500">{viewingCarBooking.phone || "-"}</p>
                </div>
                <div className="flex items-center gap-1 shrink-0">
                  <button
                    onClick={() => handlePrintCar(viewingCarBooking)}
                    className="text-stone-400 hover:text-teal-800 p-1.5"
                    title="Print"
                  >
                    <Printer size={18} />
                  </button>
                  <button
                    onClick={() => setCopyPickerSource({ type: "cars", record: viewingCarBooking })}
                    className="text-stone-400 hover:text-amber-600 p-1.5"
                    title="Copy to a file"
                  >
                    <FileText size={18} />
                  </button>
                  {carsPerm.canEdit && (
                    <button
                      onClick={() => { handleEditCarClick(viewingCarBooking); setViewingCarBooking(null); }}
                      className="text-stone-400 hover:text-teal-800 p-1.5"
                      title="Edit"
                    >
                      <Pencil size={18} />
                    </button>
                  )}
                  {carsPerm.canDelete && (
                    <button
                      onClick={() => {
                        const id = viewingCarBooking.id;
                        handleDeleteCar(id, () => setViewingCarBooking(null));
                      }}
                      className="text-stone-400 hover:text-red-600 p-1.5"
                      title="Delete"
                    >
                      <Trash2 size={18} />
                    </button>
                  )}
                  <button
                    onClick={() => setViewingCarBooking(null)}
                    className="text-stone-400 hover:text-stone-700 p-1.5"
                    title="Close"
                  >
                    <X size={20} />
                  </button>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4 text-sm">
                <div>
                  <span className="text-stone-500">Route: </span>
                  {viewingCarBooking.routeFrom || "-"} → {viewingCarBooking.routeTo || "-"}
                </div>
                <div><span className="text-stone-500">Car type: </span>{viewingCarBooking.carType || "-"}</div>
                <div><span className="text-stone-500">Supplier: </span>{viewingCarBooking.supplier || "-"}</div>
                <div>
                  <span className="text-stone-500">Trip: </span>
                  {viewingCarBooking.isRoundTrip ? "Round trip" : "One way"}
                </div>
                <div>
                  <span className="text-stone-500">Waiting: </span>
                  {viewingCarBooking.hasWaiting ? `${viewingCarBooking.waitingHours || 0} h` : "-"}
                </div>
                <div>
                  <span className="text-stone-500">Flight number: </span>
                  {viewingCarBooking.startsAtAirport ? (viewingCarBooking.flightNumber || "-") : "-"}
                </div>
                <div>
                  <span className="text-stone-500">Booking date: </span>
                  {viewingCarBooking.bookingDate ? formatDisplayDate(viewingCarBooking.bookingDate) : "-"}
                  {viewingCarBooking.bookingTime ? ` · ${viewingCarBooking.bookingTime}` : ""}
                </div>
                {viewingCarBooking.isRoundTrip && (
                  <div>
                    <span className="text-stone-500">Return: </span>
                    {viewingCarBooking.returnDate ? formatDisplayDate(viewingCarBooking.returnDate) : "-"}
                    {viewingCarBooking.returnTime ? ` · ${viewingCarBooking.returnTime}` : ""}
                  </div>
                )}
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4">
                <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                  <p className="text-[11px] text-stone-500">Collection</p>
                  <p className="text-sm font-bold text-stone-800">
                    {viewingCarBooking.collection
                      ? `${fmt(parseFloat(viewingCarBooking.collection) || 0)} ${viewingCarBooking.currency}`
                      : "-"}
                  </p>
                </div>
                <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                  <p className="text-[11px] text-stone-500">Driver tip</p>
                  <p className="text-sm font-bold text-stone-800">
                    {viewingCarBooking.driverTip
                      ? `${fmt(parseFloat(viewingCarBooking.driverTip) || 0)} ${viewingCarBooking.currency}`
                      : "-"}
                  </p>
                </div>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                  <p className="text-[11px] text-stone-500">Net</p>
                  <p className="text-sm font-bold text-stone-800">
                    {fmt(parseFloat(viewingCarBooking.netPrice) || 0)} {viewingCarBooking.currency}
                  </p>
                </div>
                <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                  <p className="text-[11px] text-stone-500">Sold</p>
                  <p className="text-sm font-bold text-stone-800">
                    {fmt(parseFloat(viewingCarBooking.soldPrice) || 0)} {viewingCarBooking.currency}
                  </p>
                </div>
                <div className="bg-stone-50 border border-stone-200 rounded-xl px-3 py-2 text-center">
                  <p className="text-[11px] text-stone-500">Profit</p>
                  <p className="text-sm font-bold text-emerald-700">
                    {fmt((parseFloat(viewingCarBooking.soldPrice) || 0) - (parseFloat(viewingCarBooking.netPrice) || 0))} {viewingCarBooking.currency}
                  </p>
                </div>
              </div>
            </div>
          </div>
        )}
        </>
        )}

        {activeSection === "files" && (
          <>
            {fileError && (
              <div className="bg-red-50 text-red-700 text-sm rounded-xl px-3 py-2 mb-4">{fileError}</div>
            )}

            {!openFile && !draftFile && (
              <>
                {/* Summary cards, same style as the Flights section */}
                <div className="grid grid-cols-3 gap-1.5 sm:gap-3 mb-6">
                  <div className="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0">
                    <div className="bg-stone-100 rounded-xl p-1.5 sm:p-2 text-stone-600 shrink-0"><FileText size={18} className="sm:hidden" /><FileText size={20} className="hidden sm:block" /></div>
                    <div className="min-w-0">
                      <p className="text-xs text-stone-500">Files</p>
                      <p className="text-sm sm:text-lg font-bold truncate">{visibleFiles.length}</p>
                    </div>
                  </div>
                  <div className="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0">
                    <div className="bg-teal-50 rounded-xl p-1.5 sm:p-2 text-teal-900 shrink-0"><Wallet size={18} className="sm:hidden" /><Wallet size={20} className="hidden sm:block" /></div>
                    <div className="min-w-0">
                      <p className="text-xs text-stone-500">Total sales</p>
                      <p className="text-sm sm:text-lg font-bold truncate">{fmt(filesGrandTotals.sold)}</p>
                    </div>
                  </div>
                  <div className="bg-white rounded-2xl border border-stone-200 p-2.5 sm:p-4 flex items-center gap-2 sm:gap-3 min-w-0">
                    <div className="bg-emerald-50 rounded-xl p-1.5 sm:p-2 text-emerald-700 shrink-0"><TrendingUp size={18} className="sm:hidden" /><TrendingUp size={20} className="hidden sm:block" /></div>
                    <div className="min-w-0">
                      <p className="text-xs text-stone-500">Total profit</p>
                      <p className="text-sm sm:text-lg font-bold text-emerald-700 truncate">{fmt(filesGrandTotals.profit)}</p>
                    </div>
                  </div>
                </div>

                {/* Search and filters — same unified card style as the other sections. */}
                <div className="bg-white border border-stone-200 rounded-2xl p-3 sm:p-4 mb-4">
                  <div className="flex items-stretch gap-2">
                    <div className="relative flex-1">
                      <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
                      <input
                        className="w-full border border-stone-300 rounded-xl pl-9 pr-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        placeholder="Search by serial, company, notes, or employee"
                        value={fileQuery}
                        onChange={(e) => setFileQuery(e.target.value)}
                      />
                    </div>
                    <button
                      type="button"
                      onClick={() => setFileFiltersOpen(!fileFiltersOpen)}
                      className={`shrink-0 flex items-center gap-1.5 border rounded-xl px-3 py-2 text-sm font-medium transition-colors ${
                        fileFiltersOpen ? "border-teal-700 text-teal-800 bg-teal-50" : "border-stone-300 text-stone-600 hover:bg-stone-50 bg-white"
                      }`}
                    >
                      <SlidersHorizontal size={16} />
                      <span className="hidden sm:inline">Filters</span>
                      {activeFileFilterCount > 0 && (
                        <span className="bg-teal-700 text-white text-[11px] font-bold rounded-full min-w-[20px] h-5 px-1 flex items-center justify-center">
                          {activeFileFilterCount}
                        </span>
                      )}
                      <ChevronDown size={14} className={`transition-transform ${fileFiltersOpen ? "rotate-180" : ""}`} />
                    </button>
                  </div>

                  {fileFiltersOpen && (
                    <div className="flex flex-wrap gap-3 mt-3 pt-3 border-t border-stone-100">
                      <div>
                        <label className="text-xs text-stone-500 block mb-1">Year</label>
                        <MultiSelectDropdown
                          label="years"
                          icon={Calendar}
                          options={fileYearsAvailable}
                          selected={fileSelectedYear}
                          onChange={setFileSelectedYear}
                          placeholder="All years"
                        />
                      </div>
                      <div>
                        <label className="text-xs text-stone-500 block mb-1">Company</label>
                        <MultiSelectDropdown
                          label="companies"
                          icon={Building2}
                          options={fileCompaniesAvailable}
                          selected={fileSelectedCompany}
                          onChange={setFileSelectedCompany}
                          placeholder="All companies"
                        />
                      </div>
                      <div>
                        <label className="text-xs text-stone-500 block mb-1">Employee</label>
                        <MultiSelectDropdown
                          label="employees"
                          icon={User}
                          options={fileEmployeesAvailable}
                          selected={fileSelectedEmployee}
                          onChange={setFileSelectedEmployee}
                          placeholder="All employees"
                        />
                      </div>
                    </div>
                  )}

                  {hasActiveFileFilter && (
                    <div className="flex flex-wrap items-center gap-x-4 gap-y-1.5 mt-3 pt-3 border-t border-stone-100 text-xs">
                      <span className="text-stone-400 font-medium">Applied:</span>
                      {fileSelectedYear.map((y) => (
                        <span key={`year-${y}`} className="inline-flex items-center gap-1 text-stone-600">
                          Year: <span className="font-semibold text-stone-800">{y}</span>
                          <button onClick={() => setFileSelectedYear(fileSelectedYear.filter((v) => v !== y))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                        </span>
                      ))}
                      {fileSelectedCompany.map((c) => (
                        <span key={`company-${c}`} className="inline-flex items-center gap-1 text-stone-600">
                          Company: <span className="font-semibold text-stone-800">{c}</span>
                          <button onClick={() => setFileSelectedCompany(fileSelectedCompany.filter((v) => v !== c))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                        </span>
                      ))}
                      {fileSelectedEmployee.map((e) => (
                        <span key={`employee-${e}`} className="inline-flex items-center gap-1 text-stone-600">
                          Employee: <span className="font-semibold text-stone-800">{e}</span>
                          <button onClick={() => setFileSelectedEmployee(fileSelectedEmployee.filter((v) => v !== e))} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                        </span>
                      ))}
                      {fileQuery.trim() && (
                        <span className="inline-flex items-center gap-1 text-stone-600">
                          Search: <span className="font-semibold text-stone-800">"{fileQuery.trim()}"</span>
                          <button onClick={() => setFileQuery("")} className="text-stone-400 hover:text-red-600"><X size={12} /></button>
                        </span>
                      )}
                      <button onClick={clearAllFileFilters} className="text-red-600 hover:text-red-700 font-semibold ml-auto">
                        Clear all
                      </button>
                    </div>
                  )}
                </div>

                <div className="bg-white rounded-2xl border border-stone-200 divide-y divide-stone-100 overflow-hidden mb-6">
                  <button
                    onClick={startNewFileDraft}
                    className="w-full flex items-center gap-2 px-4 py-3 text-teal-800 hover:bg-teal-50/50 text-sm font-semibold text-left"
                  >
                    <Plus size={16} /> Create new file
                  </button>

                  {filteredFiles.length === 0 ? (
                    <p className="text-sm text-stone-400 text-center py-10">
                      {visibleFiles.length === 0
                        ? "No files yet — create one and pull in copies from Flights, Hotels, or Visa."
                        : "No files match the current search/filters."}
                    </p>
                  ) : (
                    filteredFiles.map((f) => {
                      const t = fileTotals(f);
                      return (
                        <div
                          key={f.id}
                          className="w-full flex items-center justify-between gap-3 px-4 py-3 hover:bg-teal-50/50"
                        >
                          <button
                            onClick={() => { setOpenFileId(f.id); setEditingFileServices(false); }}
                            className="flex-1 min-w-0 flex items-center justify-between gap-3 text-left"
                          >
                            <div className="min-w-0">
                              <p className="font-semibold text-stone-900 text-sm truncate">
                                {f.serial} {f.company ? `· ${f.company}` : ""}
                              </p>
                              <p className="text-xs text-stone-400">
                                {formatDisplayDate(f.createdAt)} · {f.createdBy} · {(f.items || []).length} item{(f.items || []).length === 1 ? "" : "s"}
                              </p>
                            </div>
                            <div className="text-right shrink-0">
                              <p className="text-sm font-bold">{fmt(t.sold)}</p>
                              <p className="text-xs text-emerald-700 font-semibold">+{fmt(t.profit)}</p>
                            </div>
                          </button>
                          {filesPerm.canDelete && (
                          <button
                            onClick={() =>
                              requestConfirm(`Delete file ${f.serial}? This cannot be undone.`, async () => {
                                await deleteFile(f.id);
                                setConfirmDialog(null);
                              })
                            }
                            className="text-red-500 hover:text-red-700 p-1.5 shrink-0"
                          >
                            <Trash2 size={15} />
                          </button>
                          )}
                        </div>
                      );
                    })
                  )}
                </div>
              </>
            )}

            {draftFile && (
              <div>
                <button
                  onClick={cancelDraftFile}
                  className="mb-4 text-stone-500 hover:text-teal-800 text-sm font-semibold flex items-center gap-1.5"
                >
                  <ArrowLeft size={15} /> Cancel
                </button>

                <div className="bg-white rounded-2xl border border-stone-200 p-4 md:p-5 mb-6">
                  <p className="text-xs text-stone-400 mb-4">New file — not saved yet</p>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-4">
                    <div>
                      <label className="text-xs text-stone-500 block mb-1">Serial</label>
                      <input
                        type="text"
                        className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={draftFile.serial || ""}
                        onChange={(e) => updateDraftField("serial", e.target.value)}
                      />
                    </div>
                    <div>
                      <label className="text-xs text-stone-500 block mb-1">File date</label>
                      <input
                        type="date"
                        max={todayDateStr()}
                        className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={draftFile.createdAt || ""}
                        onChange={(e) =>
                          e.target.value && updateDraftDate(e.target.value > todayDateStr() ? todayDateStr() : e.target.value)
                        }
                      />
                    </div>
                    <div>
                      <label className="text-xs text-stone-500 block mb-1">Company</label>
                      <input
                        type="text"
                        list="file-company-list"
                        className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={draftFile.company || ""}
                        onChange={(e) => updateDraftField("company", e.target.value)}
                      />
                      <datalist id="file-company-list">
                        {suggestions.companies.map((c, i) => (
                          <option key={i} value={companyName(c)} />
                        ))}
                      </datalist>
                    </div>
                    <div className="md:col-span-2">
                      <label className="text-xs text-stone-500 block mb-1">Notes</label>
                      <input
                        type="text"
                        className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={draftFile.notes || ""}
                        onChange={(e) => updateDraftField("notes", e.target.value)}
                      />
                    </div>
                  </div>

                  <div className="flex flex-wrap gap-2 mb-4">
                    <button
                      onClick={() => setShowFilePicker(true)}
                      className="text-teal-800 border border-teal-800 hover:bg-teal-50 text-xs font-semibold rounded-xl px-3 py-2 flex items-center gap-1.5"
                    >
                      <Plus size={14} /> Add services
                    </button>
                    <button
                      onClick={confirmDraftFile}
                      className="bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-xs font-semibold rounded-xl px-3 py-2 flex items-center gap-1.5 shadow-sm shadow-teal-800/30"
                    >
                      <Plus size={14} /> Add file
                    </button>
                  </div>

                  {/* Totals for the draft's items so far */}
                  <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-4">
                    <div className="bg-stone-50 rounded-xl p-3 text-center">
                      <p className="text-[11px] text-stone-500">Net</p>
                      <p className="font-bold text-sm">{fmt(fileTotals(draftFile).net)}</p>
                    </div>
                    <div className="bg-stone-50 rounded-xl p-3 text-center">
                      <p className="text-[11px] text-stone-500">Sold</p>
                      <p className="font-bold text-sm">{fmt(fileTotals(draftFile).sold)}</p>
                    </div>
                    <div className="bg-emerald-50 rounded-xl p-3 text-center">
                      <p className="text-[11px] text-emerald-700">Profit</p>
                      <p className="font-bold text-sm text-emerald-700">{fmt(fileTotals(draftFile).profit)}</p>
                    </div>
                  </div>
                </div>

                <div className="bg-white rounded-2xl border border-stone-200 divide-y divide-stone-100 overflow-hidden">
                  {(draftFile.items || []).length === 0 ? (
                    <p className="text-sm text-stone-400 text-center py-10">No services added yet — use "Add services" above.</p>
                  ) : (
                    (draftFile.items || []).map((it) => (
                      <div key={it.id} className="flex items-center justify-between gap-3 px-4 py-3">
                        <div className="min-w-0">
                          <p className="text-xs text-teal-800 font-semibold">{FILE_SOURCE_LABELS[it.sourceType] || it.sourceType}</p>
                          <p className="text-sm text-stone-900 truncate">{it.label}</p>
                          <p className="text-xs text-stone-400">{formatDisplayDate(it.date)}</p>
                        </div>
                        <div className="flex items-center gap-3 shrink-0">
                          <div className="text-right">
                            <p className="text-sm font-bold">{fmt(it.soldPrice)} {it.currency}</p>
                            <p className="text-xs text-emerald-700">net {fmt(it.netPrice)} {it.currency}</p>
                          </div>
                          <button
                            onClick={() => removeDraftItem(it.id)}
                            className="text-red-500 hover:text-red-700 p-1"
                          >
                            <Trash2 size={15} />
                          </button>
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>
            )}

            {openFile && (
              <div>
                <button
                  onClick={() => { setOpenFileId(null); setEditingFileServices(false); }}
                  className="mb-4 text-stone-500 hover:text-teal-800 text-sm font-semibold flex items-center gap-1.5"
                >
                  <ArrowLeft size={15} /> Back to files
                </button>

                <div className="bg-white rounded-2xl border border-stone-200 p-4 md:p-5 mb-6">
                  <div className="flex items-start justify-between gap-3 mb-4">
                    <div>
                      <h2 className="font-semibold text-stone-900">{openFile.serial}</h2>
                      <p className="text-xs text-stone-400">{formatDisplayDate(openFile.createdAt)} · Created by {openFile.createdBy}</p>
                    </div>
                    <div className="flex items-center gap-2">
                      {filesPerm.canEdit && (
                      <button
                        onClick={() => setEditingFileServices((v) => !v)}
                        className={
                          editingFileServices
                            ? "bg-teal-800 text-white text-xs font-semibold rounded-xl px-3 py-1.5 flex items-center gap-1.5"
                            : "text-teal-800 border border-teal-800 hover:bg-teal-50 text-xs font-semibold rounded-xl px-3 py-1.5 flex items-center gap-1.5"
                        }
                      >
                        <Pencil size={13} /> {editingFileServices ? "Done editing" : "Edit services"}
                      </button>
                      )}
                      {filesPerm.canDelete && (
                      <button
                        onClick={() => deleteFile(openFile.id)}
                        className="text-red-600 border border-red-200 hover:bg-red-50 text-xs font-semibold rounded-xl px-3 py-1.5 flex items-center gap-1.5"
                      >
                        <Trash2 size={13} /> Delete file
                      </button>
                      )}
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-4">
                    <div>
                      <label className="text-xs text-stone-500 block mb-1">Serial</label>
                      <input
                        type="text"
                        className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={openFile.serial || ""}
                        onChange={(e) => updateFileField(openFile.id, "serial", e.target.value)}
                      />
                    </div>
                    <div>
                      <label className="text-xs text-stone-500 block mb-1">File date</label>
                      <input
                        type="date"
                        max={todayDateStr()}
                        className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={openFile.createdAt || ""}
                        onChange={(e) =>
                          e.target.value &&
                          updateFileDate(openFile.id, e.target.value > todayDateStr() ? todayDateStr() : e.target.value)
                        }
                      />
                    </div>
                    <div>
                      <label className="text-xs text-stone-500 block mb-1">Company</label>
                      <input
                        type="text"
                        list="file-company-list"
                        className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={openFile.company || ""}
                        onChange={(e) => updateFileField(openFile.id, "company", e.target.value)}
                      />
                      <datalist id="file-company-list">
                        {suggestions.companies.map((c, i) => (
                          <option key={i} value={companyName(c)} />
                        ))}
                      </datalist>
                    </div>
                    <div className="md:col-span-2">
                      <label className="text-xs text-stone-500 block mb-1">Notes</label>
                      <input
                        type="text"
                        className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
                        value={openFile.notes || ""}
                        onChange={(e) => updateFileField(openFile.id, "notes", e.target.value)}
                      />
                    </div>
                  </div>

                  {editingFileServices && (
                    <button
                      onClick={() => setShowFilePicker(true)}
                      className="mb-4 text-teal-800 border border-teal-800 hover:bg-teal-50 text-xs font-semibold rounded-xl px-3 py-2 flex items-center gap-1.5"
                    >
                      <Plus size={14} /> Add service
                    </button>
                  )}

                  {/* Totals for this file only — separate from every other section's totals */}
                  <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-4">
                    <div className="bg-stone-50 rounded-xl p-3 text-center">
                      <p className="text-[11px] text-stone-500">Net</p>
                      <p className="font-bold text-sm">{fmt(fileTotals(openFile).net)}</p>
                    </div>
                    <div className="bg-stone-50 rounded-xl p-3 text-center">
                      <p className="text-[11px] text-stone-500">Sold</p>
                      <p className="font-bold text-sm">{fmt(fileTotals(openFile).sold)}</p>
                    </div>
                    <div className="bg-emerald-50 rounded-xl p-3 text-center">
                      <p className="text-[11px] text-emerald-700">Profit</p>
                      <p className="font-bold text-sm text-emerald-700">{fmt(fileTotals(openFile).profit)}</p>
                    </div>
                  </div>
                </div>

                <div className="bg-white rounded-2xl border border-stone-200 divide-y divide-stone-100 overflow-hidden">
                  {(openFile.items || []).length === 0 ? (
                    <p className="text-sm text-stone-400 text-center py-10">No items added to this file yet.</p>
                  ) : (
                    (openFile.items || []).map((it) => (
                      <div key={it.id} className="flex items-center justify-between gap-3 px-4 py-3">
                        <div className="min-w-0">
                          <p className="text-xs text-teal-800 font-semibold">{FILE_SOURCE_LABELS[it.sourceType] || it.sourceType}</p>
                          <p className="text-sm text-stone-900 truncate">{it.label}</p>
                          <p className="text-xs text-stone-400">{formatDisplayDate(it.date)}</p>
                        </div>
                        <div className="flex items-center gap-3 shrink-0">
                          <div className="text-right">
                            <p className="text-sm font-bold">{fmt(it.soldPrice)} {it.currency}</p>
                            <p className="text-xs text-emerald-700">net {fmt(it.netPrice)} {it.currency}</p>
                          </div>
                          {editingFileServices && filesPerm.canEdit && (
                            <button
                              onClick={() => removeItemFromFile(openFile.id, it.id)}
                              className="text-red-500 hover:text-red-700 p-1"
                            >
                              <Trash2 size={15} />
                            </button>
                          )}
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>
            )}

            {/* Picker: pull a read-only copy of an existing Flights/Hotels/Visa record into
                the currently open file. Selecting a record only ever ADDS a snapshot here —
                it never edits, deletes, or otherwise affects the original record or that
                section's own totals. */}
            {showFilePicker && (openFile || draftFile) && (
              <div className="fixed inset-0 bg-stone-900/40 flex items-center justify-center p-4 z-50" onClick={() => setShowFilePicker(false)}>
                <div
                  className="bg-white rounded-2xl border border-stone-200 w-full max-w-lg max-h-[85vh] flex flex-col"
                  onClick={(ev) => ev.stopPropagation()}
                >
                  <div className="flex items-center justify-between p-4 border-b border-stone-100">
                    <h3 className="font-semibold text-stone-900">
                      {draftFile ? "Add services" : `Add a copy to ${openFile.serial}`}
                    </h3>
                    <button onClick={() => setShowFilePicker(false)} className="text-stone-400 hover:text-stone-700 p-1">
                      <X size={16} />
                    </button>
                  </div>

                  <div className="flex gap-2 px-4 pt-3">
                    {[
                      { key: "flights", label: "Flights", icon: Plane },
                      { key: "hotels", label: "Hotels", icon: Building2 },
                      { key: "visa", label: "Visa", icon: PassportIcon },
                    ].map((tab) => (
                      <button
                        key={tab.key}
                        onClick={() => setFilePickerTab(tab.key)}
                        className={`flex items-center gap-1.5 text-xs font-semibold rounded-xl px-3 py-1.5 border ${
                          filePickerTab === tab.key
                            ? "bg-teal-800 text-white border-teal-800"
                            : "bg-white text-stone-600 border-stone-300 hover:bg-stone-50"
                        }`}
                      >
                        <tab.icon size={14} className={tab.key === "flights" ? "rotate-45" : ""} /> {tab.label}
                      </button>
                    ))}
                  </div>

                  <div className="overflow-y-auto p-4 space-y-2">
                    {filePickerTab === "flights" && (
                      visibleTickets.length === 0 ? (
                        <p className="text-sm text-stone-400 text-center py-6">No tickets to add yet.</p>
                      ) : (
                        visibleTickets.map((t) => (
                          <button
                            key={t.id}
                            onClick={async () => {
                              if (draftFile) addDraftItem("flights", t);
                              else await addItemToFile(openFile.id, "flights", t);
                              setShowFilePicker(false);
                            }}
                            className="w-full text-left border border-stone-200 rounded-xl px-3 py-2 hover:bg-teal-50 hover:border-teal-300 flex items-center justify-between gap-2"
                          >
                            <span className="text-sm text-stone-800 truncate">
                              {routeLabel(t)} · {getCustomers(t).map((c) => c.name).filter(Boolean).join(", ") || "-"}
                            </span>
                            <span className="text-xs text-stone-400 shrink-0">{fmt(t.soldPrice)}</span>
                          </button>
                        ))
                      )
                    )}
                    {filePickerTab === "hotels" && (
                      visibleHotelBookings.length === 0 ? (
                        <p className="text-sm text-stone-400 text-center py-6">No hotel bookings to add yet.</p>
                      ) : (
                        visibleHotelBookings.map((h) => (
                          <button
                            key={h.id}
                            onClick={async () => {
                              if (draftFile) addDraftItem("hotels", h);
                              else await addItemToFile(openFile.id, "hotels", h);
                              setShowFilePicker(false);
                            }}
                            className="w-full text-left border border-stone-200 rounded-xl px-3 py-2 hover:bg-teal-50 hover:border-teal-300 flex items-center justify-between gap-2"
                          >
                            <span className="text-sm text-stone-800 truncate">
                              {h.hotel || "Hotel"}{h.customer ? ` · ${h.customer}` : ""}
                            </span>
                            <span className="text-xs text-stone-400 shrink-0">{fmt(hotelSoldTotal(h))}</span>
                          </button>
                        ))
                      )
                    )}
                    {filePickerTab === "visa" && (
                      visibleVisaBookingsForFiles.length === 0 ? (
                        <p className="text-sm text-stone-400 text-center py-6">No visa bookings to add yet.</p>
                      ) : (
                        visibleVisaBookingsForFiles.map((v) => (
                          <button
                            key={v.id}
                            onClick={async () => {
                              if (draftFile) addDraftItem("visa", v);
                              else await addItemToFile(openFile.id, "visa", v);
                              setShowFilePicker(false);
                            }}
                            className="w-full text-left border border-stone-200 rounded-xl px-3 py-2 hover:bg-teal-50 hover:border-teal-300 flex items-center justify-between gap-2"
                          >
                            <span className="text-sm text-stone-800 truncate">
                              {v.visaType || "Visa"} · {(v.customers || []).map((c) => c.name).filter(Boolean).join(", ") || "-"}
                            </span>
                            <span className="text-xs text-stone-400 shrink-0">{fmt(visaSoldTotal(v))} {v.currency}</span>
                          </button>
                        ))
                      )
                    )}
                  </div>
                </div>
              </div>
            )}
          </>
        )}

        {activeSection === "accounts" && canAccessAccounts && (
        <>
        {accountsError && (
          <div className="text-sm rounded-xl px-3 py-2 mb-4 bg-red-50 text-red-700">{accountsError}</div>
        )}

        {/* Accounts sub-tab switcher */}
        <div className="flex items-center gap-2 mb-5 overflow-x-auto -mx-4 px-4 md:mx-0 md:px-0">
          {[
            { key: "overview", label: at("tabOverview"), icon: PieChart },
            { key: "suppliers", label: at("tabSuppliers"), icon: Building2 },
            { key: "customers", label: at("tabCustomers"), icon: Users },
            { key: "treasury", label: at("tabTreasury"), icon: Landmark },
            { key: "expenses", label: at("tabExpenses"), icon: Receipt },
            { key: "reports", label: at("tabReports"), icon: ClipboardList },
          ].map((tab) => (
            <button
              key={tab.key}
              onClick={() => setAccountsTab(tab.key)}
              className={`shrink-0 flex items-center gap-1.5 px-3.5 py-2 rounded-xl border text-xs font-semibold transition-colors ${
                accountsTab === tab.key
                  ? "bg-teal-800 text-white border-teal-800"
                  : "bg-white text-stone-500 border-stone-200 hover:border-teal-300 hover:text-teal-800"
              }`}
            >
              <tab.icon size={15} />
              {tab.label}
            </button>
          ))}
        </div>

        {/* ---------- Overview ---------- */}
        {accountsTab === "overview" && (
          <div>
            <div className="grid grid-cols-2 md:grid-cols-3 gap-3 mb-6">
              <div className="bg-white rounded-2xl border border-stone-200 p-4">
                <p className="text-xs text-stone-500 mb-1">{at("monthRevenue")}</p>
                <p className="text-lg font-bold text-emerald-700">{fmt(monthRevenue)} {acctCurrency}</p>
              </div>
              <div className="bg-white rounded-2xl border border-stone-200 p-4">
                <p className="text-xs text-stone-500 mb-1">{at("monthExpenses")}</p>
                <p className="text-lg font-bold text-red-600">{fmt(monthExpenses)} {acctCurrency}</p>
              </div>
              <div className="bg-white rounded-2xl border border-stone-200 p-4">
                <p className="text-xs text-stone-500 mb-1">{at("monthNetProfit")}</p>
                <p className={`text-lg font-bold ${monthRevenue - monthExpenses >= 0 ? "text-emerald-700" : "text-red-600"}`}>
                  {fmt(monthRevenue - monthExpenses)} {acctCurrency}
                </p>
              </div>
              <div className="bg-white rounded-2xl border border-stone-200 p-4">
                <p className="text-xs text-stone-500 mb-1">{at("totalTreasuryBalance")}</p>
                <p className="text-lg font-bold text-teal-800">{fmt(totalTreasuryBalance)} {acctCurrency}</p>
              </div>
              <div className="bg-white rounded-2xl border border-stone-200 p-4">
                <p className="text-xs text-stone-500 mb-1">{at("totalOwedSuppliers")}</p>
                <p className="text-lg font-bold text-amber-700">{fmt(totalSupplierBalance)} {acctCurrency}</p>
              </div>
              <div className="bg-white rounded-2xl border border-stone-200 p-4">
                <p className="text-xs text-stone-500 mb-1">{at("totalOwedCustomers")}</p>
                <p className="text-lg font-bold text-amber-700">{fmt(totalCustomerBalance)} {acctCurrency}</p>
              </div>
            </div>

            <h3 className="text-sm font-bold text-stone-700 mb-2">{at("profitBySection")}</h3>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              {["flights", "hotels", "visa", "cars"].map((sec) => {
                const val =
                  sec === "flights"
                    ? tickets.filter((t) => (t.date || "").slice(0, 7) === thisMonthPrefix).reduce((s, t) => s + profitAfterRefund(t), 0)
                    : sec === "hotels"
                    ? hotelBookings.filter((h) => (h.bookingDate || "").slice(0, 7) === thisMonthPrefix).reduce((s, h) => s + hotelProfitTotal(h), 0)
                    : sec === "visa"
                    ? visaBookings.filter((v) => (v.bookingDate || "").slice(0, 7) === thisMonthPrefix).reduce((s, v) => s + hotelInEgp(visaProfitTotal(v), v.currency), 0)
                    : carBookings.filter((c) => (c.bookingDate || "").slice(0, 7) === thisMonthPrefix).reduce((s, c) => s + hotelInEgp((parseFloat(c.soldPrice) || 0) - (parseFloat(c.netPrice) || 0), c.currency), 0);
                return (
                  <div key={sec} className="bg-white rounded-2xl border border-stone-200 p-4">
                    <p className="text-xs text-stone-500 mb-1">{sectionLabel(sec)}</p>
                    <p className="text-base font-bold text-emerald-700">{fmt(val)} {acctCurrency}</p>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* ---------- Suppliers ---------- */}
        {accountsTab === "suppliers" && (
          <div>
            <div className="relative mb-4">
              <Search size={16} className="absolute right-3 top-1/2 -translate-y-1/2 text-stone-400" />
              <input
                value={supplierQuery}
                onChange={(e) => setSupplierQuery(e.target.value)}
                placeholder={at("searchSupplier")}
                className="w-full border border-stone-300 rounded-xl pr-9 pl-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
              />
            </div>
            <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden">
              <table className="w-full text-sm">
                <thead className="bg-stone-50 text-stone-500 text-xs">
                  <tr>
                    <th className="text-right px-3 py-2 font-medium">{at("colSupplier")}</th>
                    <th className="text-right px-3 py-2 font-medium">{at("colSections")}</th>
                    <th className="text-right px-3 py-2 font-medium">{at("colTotalOwed")}</th>
                    <th className="text-right px-3 py-2 font-medium">{at("colPaid")}</th>
                    <th className="text-right px-3 py-2 font-medium">{at("colRemaining")}</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-stone-100">
                  {filteredSupplierLedger.length === 0 ? (
                    <tr><td colSpan={5} className="text-center text-stone-400 py-6">{at("noSuppliers")}</td></tr>
                  ) : (
                    filteredSupplierLedger.map((s) => (
                      <tr
                        key={s.supplier}
                        onClick={() => setViewingSupplier(s.supplier)}
                        className="hover:bg-teal-50 cursor-pointer"
                      >
                        <td className="px-3 py-2 font-semibold text-stone-800">{s.supplier}</td>
                        <td className="px-3 py-2 text-stone-500 text-xs">{s.sections.map((x) => sectionLabel(x)).join(accountsLang === "en" ? ", " : "، ") || "-"}</td>
                        <td className="px-3 py-2 text-stone-700">{fmt(s.totalOwed)}</td>
                        <td className="px-3 py-2 text-emerald-700">{fmt(s.paid)}</td>
                        <td className={`px-3 py-2 font-bold ${s.balance > 0 ? "text-red-600" : "text-stone-400"}`}>{fmt(s.balance)}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ---------- Customers ---------- */}
        {accountsTab === "customers" && (
          <div>
            <div className="relative mb-4">
              <Search size={16} className="absolute right-3 top-1/2 -translate-y-1/2 text-stone-400" />
              <input
                value={customerQuery}
                onChange={(e) => setCustomerQuery(e.target.value)}
                placeholder={at("searchCustomer")}
                className="w-full border border-stone-300 rounded-xl pr-9 pl-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700"
              />
            </div>
            <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden">
              <table className="w-full text-sm">
                <thead className="bg-stone-50 text-stone-500 text-xs">
                  <tr>
                    <th className="text-right px-3 py-2 font-medium">{at("colCustomer")}</th>
                    <th className="text-right px-3 py-2 font-medium">{at("colSections")}</th>
                    <th className="text-right px-3 py-2 font-medium">{at("colTotalDue")}</th>
                    <th className="text-right px-3 py-2 font-medium">{at("colCollected")}</th>
                    <th className="text-right px-3 py-2 font-medium">{at("colRemaining")}</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-stone-100">
                  {filteredCustomerLedger.length === 0 ? (
                    <tr><td colSpan={5} className="text-center text-stone-400 py-6">{at("noCustomers")}</td></tr>
                  ) : (
                    filteredCustomerLedger.map((c) => (
                      <tr
                        key={c.customer}
                        onClick={() => setViewingCustomer(c.customer)}
                        className="hover:bg-teal-50 cursor-pointer"
                      >
                        <td className="px-3 py-2 font-semibold text-stone-800">{c.customer}</td>
                        <td className="px-3 py-2 text-stone-500 text-xs">{c.sections.map((x) => sectionLabel(x)).join(accountsLang === "en" ? ", " : "، ") || "-"}</td>
                        <td className="px-3 py-2 text-stone-700">{fmt(c.totalDue)}</td>
                        <td className="px-3 py-2 text-emerald-700">{fmt(c.paid)}</td>
                        <td className={`px-3 py-2 font-bold ${c.balance > 0 ? "text-red-600" : "text-stone-400"}`}>{fmt(c.balance)}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ---------- Treasury ---------- */}
        {accountsTab === "treasury" && (
          <div>
            <div className="flex items-center justify-between mb-3">
              <h3 className="text-sm font-bold text-stone-700">{at("accountsAndTreasuries")}</h3>
              <button
                onClick={() => { setTreasuryForm(getEmptyTreasuryAccountForm()); setTreasuryAccountEditingId(null); setShowTreasuryAccountForm(true); }}
                className="flex items-center gap-1 text-xs font-semibold text-teal-800 bg-teal-50 hover:bg-teal-100 rounded-lg px-3 py-1.5"
              >
                <Plus size={14} /> {at("addAccount")}
              </button>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 mb-6">
              {treasuryAccounts.length === 0 && (
                <p className="text-sm text-stone-400 col-span-full">{at("noAccountsYet")}</p>
              )}
              {treasuryAccounts.map((a) => (
                <div key={a.id} className="bg-white rounded-2xl border border-stone-200 p-4">
                  <div className="flex items-center justify-between mb-1">
                    <div className="flex items-center gap-1.5 text-stone-800 font-semibold text-sm">
                      {a.type === "bank" ? <Landmark size={16} /> : <Banknote size={16} />}
                      {a.name}
                    </div>
                    <div className="flex items-center gap-1">
                      <button onClick={() => handleEditTreasuryAccountClick(a)} className="text-stone-400 hover:text-teal-700"><Pencil size={14} /></button>
                      <button onClick={() => handleDeleteTreasuryAccount(a.id)} className="text-stone-400 hover:text-red-600"><Trash2 size={14} /></button>
                    </div>
                  </div>
                  <p className="text-xs text-stone-400 mb-2">{treasuryAccountTypeLabel(a.type)}</p>
                  <p className={`text-lg font-bold ${treasuryBalance(a.id) >= 0 ? "text-emerald-700" : "text-red-600"}`}>
                    {fmt(treasuryBalance(a.id))} {acctCurrency}
                  </p>
                </div>
              ))}
            </div>

            <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
              <h3 className="text-sm font-bold text-stone-700">{at("treasuryMovement")}</h3>
              <div className="flex items-center gap-2">
                <select
                  value={treasuryFilterAccountId}
                  onChange={(e) => setTreasuryFilterAccountId(e.target.value)}
                  className="border border-stone-300 rounded-lg px-2 py-1.5 text-xs"
                >
                  <option value="">{at("allAccounts")}</option>
                  {treasuryAccounts.map((a) => (
                    <option key={a.id} value={a.id}>{a.name}</option>
                  ))}
                </select>
                <button
                  onClick={() => { setTreasuryEntryForm(getEmptyTreasuryEntryForm()); setShowTreasuryEntryForm(true); }}
                  className="flex items-center gap-1 text-xs font-semibold text-teal-800 bg-teal-50 hover:bg-teal-100 rounded-lg px-3 py-1.5"
                >
                  <Plus size={14} /> {at("manualEntry")}
                </button>
              </div>
            </div>
            <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden">
              <table className="w-full text-sm">
                <thead className="bg-stone-50 text-stone-500 text-xs">
                  <tr>
                    <th className="text-right px-3 py-2 font-medium">{at("colDate")}</th>
                    <th className="text-right px-3 py-2 font-medium">{at("colAccount")}</th>
                    <th className="text-right px-3 py-2 font-medium">{at("colStatement")}</th>
                    <th className="text-right px-3 py-2 font-medium">{at("colAmount")}</th>
                    <th className="px-3 py-2"></th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-stone-100">
                  {filteredTreasuryTransactions.length === 0 ? (
                    <tr><td colSpan={5} className="text-center text-stone-400 py-6">{at("noTransactions")}</td></tr>
                  ) : (
                    filteredTreasuryTransactions.map((tx) => {
                      const [prefix, rawId] = tx.id.split(/-(.+)/);
                      return (
                        <tr key={tx.id}>
                          <td className="px-3 py-2 text-stone-500 text-xs whitespace-nowrap">{tx.date ? formatDisplayDate(tx.date) : "-"}</td>
                          <td className="px-3 py-2 text-stone-600 text-xs">{treasuryAccounts.find((a) => a.id === tx.accountId)?.name || "-"}</td>
                          <td className="px-3 py-2 text-stone-800 flex items-center gap-1.5">
                            {tx.direction === "in" ? <ArrowDownCircle size={14} className="text-emerald-600 shrink-0" /> : <ArrowUpCircle size={14} className="text-red-500 shrink-0" />}
                            {tx.label}
                          </td>
                          <td className={`px-3 py-2 font-semibold whitespace-nowrap ${tx.direction === "in" ? "text-emerald-700" : "text-red-600"}`}>
                            {tx.direction === "in" ? "+" : "-"}{fmt(tx.amount)}
                          </td>
                          <td className="px-3 py-2">
                            {prefix === "te" && (
                              <button onClick={() => handleDeleteTreasuryEntry(rawId)} className="text-stone-400 hover:text-red-600"><Trash2 size={14} /></button>
                            )}
                            {prefix === "sp" && (
                              <button onClick={() => handleDeleteSupplierPayment(rawId)} className="text-stone-400 hover:text-red-600"><Trash2 size={14} /></button>
                            )}
                            {prefix === "cp" && (
                              <button onClick={() => handleDeleteCustomerPayment(rawId)} className="text-stone-400 hover:text-red-600"><Trash2 size={14} /></button>
                            )}
                            {prefix === "ex" && (
                              <button onClick={() => handleDeleteExpense(rawId)} className="text-stone-400 hover:text-red-600"><Trash2 size={14} /></button>
                            )}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ---------- Expenses ---------- */}
        {accountsTab === "expenses" && (
          <div>
            <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
              <select
                value={expenseCategoryFilter}
                onChange={(e) => setExpenseCategoryFilter(e.target.value)}
                className="border border-stone-300 rounded-lg px-2 py-1.5 text-xs"
              >
                <option value="">{at("allCategories")}</option>
                {EXPENSE_CATEGORIES.map((c) => (
                  <option key={c} value={c}>{expenseCategoryLabel(c)}</option>
                ))}
              </select>
              <button
                onClick={() => { setExpenseForm(getEmptyExpenseForm()); setExpenseEditingId(null); setShowExpenseForm(true); }}
                className="flex items-center gap-1 text-xs font-semibold text-teal-800 bg-teal-50 hover:bg-teal-100 rounded-lg px-3 py-1.5"
              >
                <Plus size={14} /> {at("addExpense")}
              </button>
            </div>
            <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden">
              <table className="w-full text-sm">
                <thead className="bg-stone-50 text-stone-500 text-xs">
                  <tr>
                    <th className="text-right px-3 py-2 font-medium">{at("colDate")}</th>
                    <th className="text-right px-3 py-2 font-medium">{at("colCategory")}</th>
                    <th className="text-right px-3 py-2 font-medium">{at("colDescription")}</th>
                    <th className="text-right px-3 py-2 font-medium">{at("colAccount")}</th>
                    <th className="text-right px-3 py-2 font-medium">{at("colAmount")}</th>
                    <th className="px-3 py-2"></th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-stone-100">
                  {filteredExpenses.length === 0 ? (
                    <tr><td colSpan={6} className="text-center text-stone-400 py-6">{at("noExpenses")}</td></tr>
                  ) : (
                    filteredExpenses.map((e) => (
                      <tr key={e.id}>
                        <td className="px-3 py-2 text-stone-500 text-xs whitespace-nowrap">{e.date ? formatDisplayDate(e.date) : "-"}</td>
                        <td className="px-3 py-2 text-stone-700">{expenseCategoryLabel(e.category)}</td>
                        <td className="px-3 py-2 text-stone-500 text-xs">{e.description || "-"}</td>
                        <td className="px-3 py-2 text-stone-500 text-xs">{treasuryAccounts.find((a) => a.id === e.accountId)?.name || "-"}</td>
                        <td className="px-3 py-2 font-semibold text-red-600 whitespace-nowrap">{fmt(parseFloat(e.amount) || 0)}</td>
                        <td className="px-3 py-2">
                          <div className="flex items-center gap-1.5 justify-end">
                            <button onClick={() => handleEditExpenseClick(e)} className="text-stone-400 hover:text-teal-700"><Pencil size={14} /></button>
                            <button onClick={() => handleDeleteExpense(e.id)} className="text-stone-400 hover:text-red-600"><Trash2 size={14} /></button>
                          </div>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ---------- Reports ---------- */}
        {accountsTab === "reports" && (
          <div>
            <div className="flex items-center gap-2 mb-4 flex-wrap">
              {[
                { key: "today", label: at("rangeToday") },
                { key: "month", label: at("rangeMonth") },
                { key: "custom", label: at("rangeCustom") },
              ].map((r) => (
                <button
                  key={r.key}
                  onClick={() => setReportsRange(r.key)}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold border ${
                    reportsRange === r.key ? "bg-teal-800 text-white border-teal-800" : "bg-white text-stone-500 border-stone-200"
                  }`}
                >
                  {r.label}
                </button>
              ))}
              {reportsRange === "custom" && (
                <>
                  <input type="date" value={reportsFrom} onChange={(e) => setReportsFrom(e.target.value)} className="border border-stone-300 rounded-lg px-2 py-1.5 text-xs" />
                  <span className="text-xs text-stone-400">{at("to")}</span>
                  <input type="date" value={reportsTo} onChange={(e) => setReportsTo(e.target.value)} className="border border-stone-300 rounded-lg px-2 py-1.5 text-xs" />
                </>
              )}
              <button
                onClick={handleExportAccountsReport}
                className="mr-auto flex items-center gap-1.5 text-xs font-semibold text-white bg-gradient-to-b from-teal-700 to-teal-900 rounded-lg px-3 py-1.5"
              >
                <Download size={14} /> {at("exportExcel")}
              </button>
            </div>

            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
              {["flights", "hotels", "visa", "cars"].map((sec) => (
                <div key={sec} className="bg-white rounded-2xl border border-stone-200 p-4">
                  <p className="text-xs text-stone-500 mb-1">{at("revenueOf")(sectionLabel(sec))}</p>
                  <p className="text-base font-bold text-emerald-700">{fmt(reportRevenueBySection[sec])} {acctCurrency}</p>
                  <p className="text-[11px] text-stone-400 mt-0.5">{reportBookingsCount[sec]} {at("bookingsCount")}</p>
                </div>
              ))}
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mb-6">
              <div className="bg-white rounded-2xl border border-stone-200 p-4">
                <p className="text-xs text-stone-500 mb-1">{at("totalRevenue")}</p>
                <p className="text-lg font-bold text-emerald-700">{fmt(reportTotalRevenue)} {acctCurrency}</p>
              </div>
              <div className="bg-white rounded-2xl border border-stone-200 p-4">
                <p className="text-xs text-stone-500 mb-1">{at("totalExpenses")}</p>
                <p className="text-lg font-bold text-red-600">{fmt(reportTotalExpenses)} {acctCurrency}</p>
              </div>
              <div className="bg-white rounded-2xl border border-stone-200 p-4">
                <p className="text-xs text-stone-500 mb-1">{at("netProfit")}</p>
                <p className={`text-lg font-bold ${reportNetProfit >= 0 ? "text-emerald-700" : "text-red-600"}`}>{fmt(reportNetProfit)} {acctCurrency}</p>
              </div>
            </div>

            <h3 className="text-sm font-bold text-stone-700 mb-2">{at("expensesByCategory")}</h3>
            <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden">
              <table className="w-full text-sm">
                <thead className="bg-stone-50 text-stone-500 text-xs">
                  <tr>
                    <th className="text-right px-3 py-2 font-medium">{at("colCategory")}</th>
                    <th className="text-right px-3 py-2 font-medium">{at("colAmount")}</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-stone-100">
                  {Object.keys(reportExpensesByCategory).length === 0 ? (
                    <tr><td colSpan={2} className="text-center text-stone-400 py-6">{at("noExpensesInPeriod")}</td></tr>
                  ) : (
                    Object.entries(reportExpensesByCategory).map(([cat, amt]) => (
                      <tr key={cat}>
                        <td className="px-3 py-2 text-stone-700">{expenseCategoryLabel(cat)}</td>
                        <td className="px-3 py-2 font-semibold text-red-600">{fmt(amt)}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}
        </>
        )}
        </>
        ) : (
          <div className="bg-white rounded-2xl border border-stone-200 p-8 text-center max-w-sm mx-auto">
            <Lock size={24} className="text-stone-300 mx-auto mb-3" />
            <h2 className="font-semibold text-stone-900 mb-1">App not activated</h2>
            <p className="text-xs text-stone-500 mb-4">
              {currentUser.isAdmin
                ? "This app isn't activated yet. Click \"Activate license\" above to enter an activation code."
                : "This app hasn't been activated yet. Please contact your admin to activate it."}
            </p>
            {currentUser.isAdmin && (
              <button
                onClick={() => setShowLicensePanel(true)}
                className="bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 shadow-sm shadow-teal-800/30 ring-1 ring-inset ring-white/10 transition-colors"
              >
                Activate license
              </button>
            )}
          </div>
        )}
        </>
        )}
      </div>
      {hasAdminAccess && (
        <div className="hidden xl:block w-64 shrink-0 pt-4 md:pt-6">
          {showOnlineList && (
          <div className="sticky top-4 flex flex-col max-h-[calc(100vh-2rem)] bg-white border border-stone-200 rounded-2xl p-2 shadow-lg shadow-stone-900/5 overflow-y-auto z-30 anim-slide-up">
            <div className="flex items-center justify-between px-1 pb-1 mb-1 border-b border-stone-100">
              <p className="text-xs font-semibold text-stone-600">{visibleOnlineUsernames.length} online now</p>
              <button onClick={() => setShowOnlineList(false)} className="text-stone-400 hover:text-stone-700 p-0.5">
                <X size={14} />
              </button>
            </div>
            {employeeRoster.length === 0 ? (
              <p className="text-[11px] text-stone-400 px-1">No employees yet</p>
            ) : (
              <ul className="space-y-0.5">
                {employeeRoster.map((e) => {
                  const online = onlineUsernames.includes(e.username);
                  const activity = online && presenceMap[e.username] && presenceMap[e.username].activity;
                  return (
                    <li key={e.username} className="flex items-center gap-1.5 px-1 py-1 text-[11px] text-stone-600">
                      <span className={`w-1.5 h-1.5 rounded-full shrink-0 ${online ? "bg-emerald-500" : "bg-stone-300"}`} />
                      <span className="flex-1 min-w-0 truncate">
                        {e.name}
                        {activity && <span className="block text-[10px] text-stone-400 truncate">{activity}</span>}
                      </span>
                      {online && (
                        <button
                          type="button"
                          onClick={() => {
                            requestConfirm(`Sign out ${e.name} now?`, () => {
                              handleForceSignOut(e.username);
                            });
                          }}
                          title="Sign out this employee"
                          className="shrink-0 inline-flex items-center gap-0.5 text-[10px] font-semibold text-red-600 hover:text-red-800 border border-red-200 hover:border-red-300 bg-red-50 rounded-full px-1.5 py-0.5"
                        >
                          <LogOut size={10} /> Sign out
                        </button>
                      )}
                    </li>
                  );
                })}
              </ul>
            )}
          </div>
          )}
        </div>
      )}
      </div>

      {activeSection === "flights" && viewingTicket && (
        <div className="fixed inset-0 bg-white z-40 overflow-y-auto">
          <div className="max-w-3xl mx-auto p-4 md:p-6">
            <div className="flex items-center justify-between mb-6">
              <div className="flex items-center gap-2">
                <div className="bg-gradient-to-b from-teal-700 to-teal-900 text-white rounded-xl p-2 shadow-sm shadow-teal-800/30">
                  <Ticket size={18} />
                </div>
                <h1 className="text-lg md:text-xl font-bold text-stone-900" style={{ fontFamily: "'Fraunces', serif" }}>Ticket details</h1>
              </div>
              <button
                onClick={closeTicketDetail}
                className="border border-stone-300 text-stone-600 text-sm rounded-xl px-3 py-2 flex items-center gap-1.5"
              >
                <X size={15} /> Close
              </button>
            </div>

            <div className="flex flex-wrap items-center gap-2 mb-4">
              <button
                onClick={() => handlePrintTicket(viewingTicket)}
                className="border border-stone-300 text-stone-600 hover:text-teal-800 hover:border-teal-700 text-sm font-semibold rounded-xl px-3 py-2 flex items-center gap-1.5"
              >
                <Printer size={15} /> Print
              </button>
              <button
                onClick={() => setCopyPickerSource({ type: "flights", record: viewingTicket })}
                className="border border-stone-300 text-stone-600 hover:text-amber-700 hover:border-amber-400 text-sm font-semibold rounded-xl px-3 py-2 flex items-center gap-1.5"
              >
                <FileText size={15} /> Copy to a file
              </button>
              {(currentUser.isAdmin || canEditTickets) && (
                <button
                  onClick={() => { handleEdit(viewingTicket); closeTicketDetail(); }}
                  className="border border-stone-300 text-stone-600 hover:text-teal-800 hover:border-teal-700 text-sm font-semibold rounded-xl px-3 py-2 flex items-center gap-1.5"
                >
                  <Pencil size={15} /> Edit
                </button>
              )}
              {(currentUser.isAdmin || canDeleteTickets) && (
                <button
                  onClick={() => { handleDelete(viewingTicket.id); closeTicketDetail(); }}
                  className="border border-stone-300 text-red-600 hover:text-red-700 hover:border-red-400 text-sm font-semibold rounded-xl px-3 py-2 flex items-center gap-1.5"
                >
                  <Trash2 size={15} /> Delete
                </button>
              )}
            </div>

            <div className="bg-white border border-stone-200 rounded-2xl divide-y divide-stone-100">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 p-4 md:p-5">
                <div>
                  <p className="text-xs text-stone-400 mb-1">Entered by</p>
                  <p className="text-sm font-medium text-stone-800">{viewingTicket.employee || "-"}</p>
                </div>
                <div>
                  <p className="text-xs text-stone-400 mb-1">Company</p>
                  <p className="text-sm font-medium text-stone-800">
                    {viewingTicket.company && viewingTicket.company.trim() ? (
                      <>{viewingTicket.company} <span className="text-teal-700 font-semibold">(Corporate)</span></>
                    ) : (
                      <span className="text-stone-400 italic">Individual</span>
                    )}
                  </p>
                </div>
                <div>
                  <p className="text-xs text-stone-400 mb-1">Supplier</p>
                  <p className="text-sm font-medium text-stone-800">{viewingTicket.supplier || "-"}</p>
                </div>
                <div>
                  <p className="text-xs text-stone-400 mb-1">Route</p>
                  <p className="text-sm font-medium text-stone-800">{routeLabel(viewingTicket)}</p>
                </div>
                <div>
                  <p className="text-xs text-stone-400 mb-1">Airline</p>
                  <p className="text-sm font-medium text-stone-800" title={getAirlineNameByIata(viewingTicket.airline) || viewingTicket.airline || ""}>
                    {viewingTicket.airline ? (getAirlineIata(viewingTicket.airline) || viewingTicket.airline) : "-"}
                  </p>
                </div>
                <div>
                  <p className="text-xs text-stone-400 mb-1">Ticket issue date</p>
                  <p className="text-sm font-medium text-stone-800">
                    {viewingTicket.date ? formatDisplayDate(viewingTicket.date) : "-"}
                  </p>
                </div>
                {viewingTicket.isReissued && (
                  <div className="sm:col-span-2 bg-amber-50 border border-amber-200 rounded-xl px-3 py-2">
                    <p className="text-xs font-semibold text-amber-800 mb-1">Exchanged ticket</p>
                    <p className="text-sm text-amber-900">
                      Old ticket number: {viewingTicket.oldTicketNumber || "-"}
                      {" · "}
                      Old issue date:{" "}
                      {viewingTicket.oldTicketIssueDate
                        ? formatDisplayDate(viewingTicket.oldTicketIssueDate)
                        : "not found"}
                    </p>
                  </div>
                )}
              </div>

              <div className="p-4 md:p-5">
                <p className="text-xs text-stone-400 mb-2">
                  Customers ({getCustomers(viewingTicket).length})
                </p>
                <div className="border border-stone-200 rounded-xl overflow-hidden">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="bg-stone-50 text-stone-500 text-xs">
                        <th className="text-left px-3 py-2 font-medium">Customer</th>
                        <th className="text-left px-3 py-2 font-medium">Ticket number</th>
                        <th className="text-left px-3 py-2 font-medium">PNR</th>
                      </tr>
                    </thead>
                    <tbody>
                      {getCustomers(viewingTicket).map((c, i) => (
                        <tr key={i} className="border-t border-stone-100">
                          <td className="px-3 py-2 text-stone-700">
                            {c.name || "-"}
                            {refundForIndex(viewingTicket, i) && (
                              <span className="ml-2 inline-block text-[10px] font-semibold text-sky-700 bg-sky-100 rounded-full px-2 py-0.5 align-middle">
                                Refunded
                              </span>
                            )}
                          </td>
                          <td className="px-3 py-2 text-stone-700 font-mono">
                            {c.ticketNumber || "-"}
                            {c.conjunction && c.ticketNumber2 && (
                              <span className="text-stone-400">{c.ticketNumber2}</span>
                            )}
                          </td>
                          <td className="px-3 py-2 text-stone-700 font-mono">{c.pnrReference || "-"}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-3 gap-4 p-4 md:p-5">
                <div>
                  <p className="text-xs text-stone-400 mb-1">Net price</p>
                  <p className="text-sm font-medium text-stone-800">{fmt(viewingTicket.netPrice)}</p>
                  {hasRefund(viewingTicket) && (
                    <p className="text-[11px] text-sky-600">After refund: {fmt(netAfterRefund(viewingTicket))}</p>
                  )}
                </div>
                <div>
                  <p className="text-xs text-stone-400 mb-1">Sold price</p>
                  <p className="text-sm font-medium text-stone-800">{fmt(viewingTicket.soldPrice)}</p>
                  {hasRefund(viewingTicket) && (
                    <p className="text-[11px] text-sky-600">After refund: {fmt(soldAfterRefund(viewingTicket))}</p>
                  )}
                </div>
                <div>
                  <p className="text-xs text-stone-400 mb-1">Profit</p>
                  <p className="text-sm font-semibold text-emerald-700">
                    {fmt(profit(viewingTicket.netPrice, viewingTicket.soldPrice))}
                  </p>
                  {hasRefund(viewingTicket) && (
                    <p className="text-[11px] text-sky-600">
                      After refund: {fmt(profitAfterRefund(viewingTicket))}
                    </p>
                  )}
                </div>
              </div>

              {/* Refund(s): entered via the checkbox box in the main ticket form (next to
                  Reissue), so this is a read-only summary — edit it by editing the
                  ticket itself. A booking with several refunded customers gets one box
                  per refund, each labeled with the customer it applies to. */}
              {hasRefund(viewingTicket) && (
                <div className="p-4 md:p-5 space-y-2">
                  {getRefunds(viewingTicket)
                    .filter((r) => r && (r.airlineAmount !== "" || r.customerAmount !== ""))
                    .map((refund, ri) => (
                      <div key={ri} className="bg-sky-50 border border-sky-200 rounded-xl px-3 py-2 flex flex-wrap gap-4 text-sm">
                        {getCustomers(viewingTicket).length > 1 && (
                          <span className="w-full">
                            <span className="text-xs text-sky-500 block">Refunded ticket</span>
                            <span className="text-sky-900 font-medium">
                              {(() => {
                                const idx = refund.customerIndex || 0;
                                const c = getCustomers(viewingTicket)[idx];
                                return c ? (c.name || `Customer ${idx + 1}`) + (c.ticketNumber ? ` — ${c.ticketNumber}` : "") : `Customer ${idx + 1}`;
                              })()}
                            </span>
                          </span>
                        )}
                        <span>
                          <span className="text-xs text-sky-500 block">Refunded by airline</span>
                          <span className="text-sky-900 font-medium">{fmt(refund.airlineAmount)}</span>
                        </span>
                        <span>
                          <span className="text-xs text-sky-500 block">Refunded to customer</span>
                          <span className="text-sky-900 font-medium">{fmt(refund.customerAmount)}</span>
                        </span>
                        {refund.date && (
                          <span>
                            <span className="text-xs text-sky-500 block">Refund date</span>
                            <span className="text-sky-900 font-medium">{formatDisplayDate(refund.date)}</span>
                          </span>
                        )}
                      </div>
                    ))}
                </div>
              )}

              <div className="p-4 md:p-5">
                <p className="text-xs text-stone-400 mb-2">Notes</p>
                <textarea
                  className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 min-h-[100px]"
                  value={notesDraft}
                  onChange={(e) => { setNotesDraft(e.target.value.toUpperCase()); setNotesSaved(false); }}
                  placeholder="No notes yet — add some here"
                />
                <div className="flex items-center gap-2 mt-3">
                  <button
                    onClick={() => saveTicketNotes(viewingTicket.id)}
                    disabled={notesDraft === (viewingTicket.notes || "")}
                    className={`text-sm font-semibold rounded-xl px-4 py-2 flex items-center gap-1.5 transition-colors ${
                      notesDraft === (viewingTicket.notes || "")
                        ? "bg-stone-200 text-stone-400 cursor-not-allowed"
                        : "bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white shadow-sm shadow-teal-800/30 ring-1 ring-inset ring-white/10"
                    }`}
                  >
                    <Check size={15} /> Save notes
                  </button>
                  {notesSaved && (
                    <span className="text-xs text-emerald-700 font-medium">Saved</span>
                  )}
                </div>

                {Array.isArray(viewingTicket.notesHistory) && viewingTicket.notesHistory.length > 0 && (
                  <div className="mt-4 pt-3 border-t border-stone-100">
                    <p className="text-xs text-stone-400 mb-2">Edit history (most recent first)</p>
                    <div className="space-y-1.5 max-h-48 overflow-y-auto">
                      {[...viewingTicket.notesHistory].reverse().map((h, idx) => (
                        <div
                          key={idx}
                          className="text-xs bg-stone-50 border border-stone-100 rounded-xl px-2.5 py-1.5 flex items-start justify-between gap-3"
                        >
                          {h.type === "edit" ? (
                            <span className="text-stone-600 break-words">
                              <span className="font-semibold text-stone-700">Ticket edited: </span>
                              {(h.changes || []).join("; ")}
                            </span>
                          ) : (
                            <span className="text-stone-600 break-words">{h.value || "(cleared)"}</span>
                          )}
                          <span className="text-stone-400 whitespace-nowrap shrink-0">
                            {h.by} · {formatDateTime(h.at)}
                          </span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {confirmDialog && (
        <div className="fixed inset-0 bg-stone-900/40 flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl border border-stone-200 p-5 w-full max-w-sm">
            <p className="text-sm text-stone-700 mb-4">{confirmDialog.message}</p>
            <div className="flex justify-end gap-2">
              <button
                onClick={() => setConfirmDialog(null)}
                className="border border-stone-300 text-stone-600 text-sm rounded-xl px-3 py-2"
              >
                Cancel
              </button>
              <button
                onClick={() => confirmDialog.onConfirm()}
                className="bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-3 py-2 shadow-sm shadow-teal-800/30 transition-colors"
              >
                Confirm
              </button>
            </div>
          </div>
        </div>
      )}

      {showRequestsPanel && currentUser && (
        <div className="fixed inset-0 bg-stone-900/40 flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-2xl border border-stone-200 p-5 w-full max-w-lg max-h-[85vh] overflow-y-auto">
            <div className="flex items-center justify-between mb-4">
              <h3 className="font-semibold text-stone-900 flex items-center gap-1.5">
                <Bell size={16} className="text-teal-800" /> Requests
              </h3>
              <button onClick={() => setShowRequestsPanel(false)} className="text-stone-400 hover:text-stone-700 p-1">
                <X size={16} />
              </button>
            </div>

            {/* Compose a new request */}
            <div className="border border-stone-200 rounded-xl p-3 mb-4 space-y-2">
              <p className="text-xs text-stone-500">New request</p>
              <select
                className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 bg-white"
                value={newRequestTo}
                onChange={(e) => setNewRequestTo(e.target.value)}
              >
                <option value="">Select an employee…</option>
                {(employees || [])
                  .filter((e) => e.username !== currentUser.username)
                  .map((e) => (
                    <option key={e.username} value={e.username}>{e.name}</option>
                  ))}
              </select>
              <textarea
                className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-700 resize-none"
                rows={2}
                placeholder="What do you need from them?"
                value={newRequestMessage}
                onChange={(e) => setNewRequestMessage(e.target.value)}
              />
              {requestSendError && <p className="text-xs text-red-600">{requestSendError}</p>}
              <button
                onClick={handleSendRequest}
                className="w-full bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-3 py-2 shadow-sm shadow-teal-800/30 transition-colors flex items-center justify-center gap-1.5"
              >
                <Send size={14} /> Send request
              </button>
            </div>

            {/* Incoming requests (addressed to me) */}
            <p className="text-xs text-stone-500 mb-1">Sent to you</p>
            <div className="space-y-2 mb-4">
              {(requests || []).filter((r) => r.toUsername === currentUser.username).length === 0 ? (
                <p className="text-xs text-stone-400">No requests yet</p>
              ) : (
                (requests || [])
                  .filter((r) => r.toUsername === currentUser.username)
                  .map((r) => (
                    <div key={r.id} className="border border-stone-200 rounded-xl p-3">
                      <div className="flex items-start justify-between gap-2">
                        <div>
                          <p className="text-sm text-stone-800"><span className="font-semibold">{r.fromName}</span></p>
                          <p className="text-sm text-stone-600 mt-0.5">{r.message}</p>
                        </div>
                        <span className={`shrink-0 text-[10px] font-semibold rounded-full px-2 py-0.5 ${
                          r.status === "pending" ? "text-amber-700 bg-amber-50 border border-amber-200" :
                          r.status === "completed" ? "text-emerald-700 bg-emerald-50 border border-emerald-200" :
                          "text-stone-500 bg-stone-100 border border-stone-200"
                        }`}>
                          {r.status === "pending" ? "Pending" : r.status === "completed" ? "Done" : "Declined"}
                        </span>
                      </div>
                      {r.status === "pending" && (
                        <div className="flex gap-2 mt-2">
                          <button
                            onClick={() => handleRespondToRequest(r.id, "completed")}
                            className="text-xs font-semibold text-emerald-700 border border-emerald-200 bg-emerald-50 hover:bg-emerald-100 rounded-full px-2.5 py-1"
                          >
                            Mark done
                          </button>
                          <button
                            onClick={() => handleRespondToRequest(r.id, "declined")}
                            className="text-xs font-semibold text-stone-500 border border-stone-200 bg-stone-50 hover:bg-stone-100 rounded-full px-2.5 py-1"
                          >
                            Decline
                          </button>
                        </div>
                      )}
                    </div>
                  ))
              )}
            </div>

            {/* Outgoing requests (sent by me) */}
            <p className="text-xs text-stone-500 mb-1">Sent by you</p>
            <div className="space-y-2">
              {(requests || []).filter((r) => r.fromUsername === currentUser.username).length === 0 ? (
                <p className="text-xs text-stone-400">You haven't sent any requests</p>
              ) : (
                (requests || [])
                  .filter((r) => r.fromUsername === currentUser.username)
                  .map((r) => (
                    <div key={r.id} className="border border-stone-200 rounded-xl p-3">
                      <div className="flex items-start justify-between gap-2">
                        <div>
                          <p className="text-sm text-stone-800">To <span className="font-semibold">{r.toName}</span></p>
                          <p className="text-sm text-stone-600 mt-0.5">{r.message}</p>
                        </div>
                        <span className={`shrink-0 text-[10px] font-semibold rounded-full px-2 py-0.5 ${
                          r.status === "pending" ? "text-amber-700 bg-amber-50 border border-amber-200" :
                          r.status === "completed" ? "text-emerald-700 bg-emerald-50 border border-emerald-200" :
                          "text-stone-500 bg-stone-100 border border-stone-200"
                        }`}>
                          {r.status === "pending" ? "Pending" : r.status === "completed" ? "Done" : "Declined"}
                        </span>
                      </div>
                    </div>
                  ))
              )}
            </div>
          </div>
        </div>
      )}

      {/* Incoming-request notification popup — floats in the corner without blocking the
          rest of the app, so it works as a lightweight "you've got a new request" alert. */}
      {incomingRequestPopup && (
        <div className="fixed top-4 right-4 z-50 w-full max-w-xs bg-white rounded-2xl border border-teal-200 shadow-xl shadow-black/10 p-4">
          <div className="flex items-start justify-between gap-2 mb-1">
            <p className="text-sm font-semibold text-stone-900 flex items-center gap-1.5">
              <Bell size={14} className="text-teal-800" /> New request
            </p>
            <button onClick={() => setIncomingRequestPopup(null)} className="text-stone-400 hover:text-stone-700 p-0.5">
              <X size={14} />
            </button>
          </div>
          <p className="text-xs text-stone-500 mb-1">From {incomingRequestPopup.fromName}</p>
          <p className="text-sm text-stone-700 mb-3">{incomingRequestPopup.message}</p>
          <div className="flex gap-2">
            <button
              onClick={() => {
                handleRespondToRequest(incomingRequestPopup.id, "completed");
                setIncomingRequestPopup(null);
              }}
              className="flex-1 text-xs font-semibold text-emerald-700 border border-emerald-200 bg-emerald-50 hover:bg-emerald-100 rounded-full px-2.5 py-1.5"
            >
              Mark done
            </button>
            <button
              onClick={() => {
                setShowRequestsPanel(true);
                setIncomingRequestPopup(null);
              }}
              className="flex-1 text-xs font-semibold text-teal-800 border border-teal-200 bg-teal-50 hover:bg-teal-100 rounded-full px-2.5 py-1.5"
            >
              Open
            </button>
          </div>
        </div>
      )}

      {openPermissionsFor && (
        <EmployeePermissionsModal
          emp={(employees || []).find((e) => e.username === openPermissionsFor && !e.isAdmin)}
          onClose={() => setOpenPermissionsFor(null)}
          onSetRole={(role) => handleRoleChange(openPermissionsFor, role)}
          onSetPermission={(field, value) => handleTogglePermission(openPermissionsFor, field, value)}
          onSetSection={(section, value) => handleToggleSection(openPermissionsFor, section, value)}
          onSetSectionPerm={(section, field, value) => handleToggleSectionPermission(openPermissionsFor, section, field, value)}
          onSave={async (draft) => {
            const err = await handleSaveEmployeeDetails(openPermissionsFor, draft);
            if (!err && draft.username !== openPermissionsFor) {
              setOpenPermissionsFor(draft.username);
            }
            return err;
          }}
          onDelete={() => {
            const target = (employees || []).find((e) => e.username === openPermissionsFor);
            setOpenPermissionsFor(null);
            if (!target) return;
            requestConfirm(`Delete "${target.name}"? This cannot be undone.`, async () => {
              await handleDeleteEmployee(target.username);
              setConfirmDialog(null);
            });
          }}
        />
      )}

      {showIataHistory && (
        <div className="fixed inset-0 bg-stone-900/40 flex items-center justify-center p-4 z-50" onClick={() => setShowIataHistory(false)}>
          <div
            className="bg-white rounded-2xl border border-stone-200 p-5 w-full max-w-sm max-h-[80vh] flex flex-col"
            onClick={(ev) => ev.stopPropagation()}
          >
            <div className="flex items-center justify-between mb-1">
              <h3 className="font-semibold text-stone-900">IATA deduction history</h3>
              <button onClick={() => setShowIataHistory(false)} className="text-stone-400 hover:text-stone-700 p-1">
                <X size={16} />
              </button>
            </div>
            <p className="text-xs text-stone-400 mb-3">Today only — resets empty at the start of each day</p>
            <div className="flex-1 overflow-y-auto -mx-1 px-1">
              {(!iataHistory || !iataHistory.deductions || iataHistory.deductions.length === 0) ? (
                <p className="text-sm text-stone-400 text-center py-6">No deductions yet today</p>
              ) : (
                <div className="border border-stone-200 rounded-xl overflow-hidden">
                  <div className="flex items-center justify-between bg-stone-50 px-3 py-2">
                    <span className="text-xs font-semibold text-stone-600">{iataHistory.date}</span>
                    <span className="text-xs font-semibold text-red-600">
                      - {fmt(iataHistory.deductions.reduce((sum, d) => sum + d.amount, 0))}
                    </span>
                  </div>
                  <div className="divide-y divide-stone-100">
                    {iataHistory.deductions.map((d, idx) => (
                      <div key={idx} className="flex items-center justify-between px-3 py-1.5 text-xs gap-2">
                        <span className="text-stone-400 shrink-0">
                          {new Date(d.time).toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit" })}
                        </span>
                        <span className="text-stone-500">{fmt(d.balanceBefore)}</span>
                        <span className="text-stone-400">→</span>
                        <span className="text-stone-600 font-semibold">{fmt(d.balanceAfter)}</span>
                        <span className="text-red-600 shrink-0">- {fmt(d.amount)}</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {copyPickerSource && (
        <div className="fixed inset-0 bg-stone-900/40 flex items-center justify-center p-4 z-50" onClick={() => setCopyPickerSource(null)}>
          <div
            className="bg-white rounded-2xl border border-stone-200 p-5 w-full max-w-sm max-h-[80vh] flex flex-col"
            onClick={(ev) => ev.stopPropagation()}
          >
            <div className="flex items-center justify-between mb-1">
              <h3 className="font-semibold text-stone-900">Copy to which file?</h3>
              <button onClick={() => setCopyPickerSource(null)} className="text-stone-400 hover:text-stone-700 p-1">
                <X size={16} />
              </button>
            </div>
            <p className="text-xs text-stone-400 mb-3">
              Adds a copy of this {FILE_SOURCE_LABELS[copyPickerSource.type] || copyPickerSource.type} record's price — the original stays untouched.
            </p>

            <button
              onClick={createFileAndCopySource}
              className="mb-3 bg-gradient-to-b from-teal-700 to-teal-900 hover:from-teal-600 hover:to-teal-800 text-white text-sm font-semibold rounded-xl px-4 py-2 shadow-sm shadow-teal-800/30 flex items-center justify-center gap-2"
            >
              <Plus size={15} /> New file (auto serial number)
            </button>

            <p className="text-xs text-stone-500 mb-1.5">Or an existing file</p>
            <div className="border border-stone-200 rounded-xl divide-y divide-stone-100 overflow-y-auto">
              {visibleFiles.length === 0 ? (
                <p className="text-xs text-stone-400 text-center py-4">No existing files yet.</p>
              ) : (
                visibleFiles.map((f) => (
                  <button
                    key={f.id}
                    onClick={() => copySourceToFile(f.id)}
                    className="w-full text-left px-3 py-2 hover:bg-teal-50 text-sm flex items-center justify-between gap-2"
                  >
                    <span className="truncate">
                      {f.serial} {f.company ? `· ${f.company}` : ""}
                    </span>
                    <span className="text-xs text-stone-400 shrink-0">{(f.items || []).length} item{(f.items || []).length === 1 ? "" : "s"}</span>
                  </button>
                ))
              )}
            </div>
          </div>
        </div>
      )}

      {/* Print preview popup — used by every service's Print button. Renders the receipt
          into an iframe inside the app instead of opening a separate browser tab, so it
          can't be blocked by a popup blocker and always looks like part of the app. */}
      {printPreview && (
        <div
          className="fixed inset-0 bg-stone-900/40 flex items-center justify-center p-4 z-50"
          onClick={() => setPrintPreview(null)}
        >
          <div
            className="bg-white rounded-2xl border border-stone-200 w-full max-w-2xl max-h-[90vh] flex flex-col overflow-hidden"
            onClick={(ev) => ev.stopPropagation()}
          >
            <div className="flex items-center justify-between px-5 py-3 border-b border-stone-100 shrink-0">
              <h3 className="text-sm font-bold text-stone-700">Print preview</h3>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => printIframeRef.current && printIframeRef.current.contentWindow.print()}
                  className="bg-gradient-to-b from-teal-700 to-teal-900 hover:brightness-110 text-white text-xs font-semibold rounded-xl px-3 py-1.5 inline-flex items-center gap-1.5"
                >
                  <Printer size={13} /> Print
                </button>
                <button
                  onClick={() => setPrintPreview(null)}
                  className="text-stone-400 hover:text-stone-700 p-1.5"
                  title="Close"
                >
                  <X size={16} />
                </button>
              </div>
            </div>
            <iframe
              ref={printIframeRef}
              title={printPreview.title}
              srcDoc={printPreview.html}
              className="flex-1 w-full"
              style={{ border: "none", minHeight: "60vh" }}
            />
          </div>
        </div>
      )}

      {/* ---------- Accounts: expense form modal ---------- */}
      {showExpenseForm && (
        <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4" onClick={() => setShowExpenseForm(false)}>
          <div className="bg-white rounded-2xl w-full max-w-md p-5" onClick={(ev) => ev.stopPropagation()}>
            <div className="flex items-center justify-between mb-4">
              <h3 className="text-sm font-bold text-stone-800">{expenseEditingId ? at("editExpense") : at("addExpense")}</h3>
              <button onClick={() => setShowExpenseForm(false)} className="text-stone-400 hover:text-stone-700"><X size={18} /></button>
            </div>
            <div className="space-y-3">
              <div>
                <label className="text-xs text-stone-500 block mb-1">{at("colDate")}</label>
                <input type="date" value={expenseForm.date} onChange={(e) => setExpenseForm({ ...expenseForm, date: e.target.value })} className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm" />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">{at("colCategory")}</label>
                <select value={expenseForm.category} onChange={(e) => setExpenseForm({ ...expenseForm, category: e.target.value })} className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm">
                  {EXPENSE_CATEGORIES.map((c) => <option key={c} value={c}>{expenseCategoryLabel(c)}</option>)}
                </select>
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">{at("descriptionOptional")}</label>
                <input value={expenseForm.description} onChange={(e) => setExpenseForm({ ...expenseForm, description: e.target.value })} className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm" />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">{at("amountEgp")}</label>
                <input type="number" value={expenseForm.amount} onChange={(e) => setExpenseForm({ ...expenseForm, amount: e.target.value })} className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm" />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">{at("payFromAccount")}</label>
                <select value={expenseForm.accountId} onChange={(e) => setExpenseForm({ ...expenseForm, accountId: e.target.value })} className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm">
                  <option value="">{at("selectAccount")}</option>
                  {treasuryAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
                </select>
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">{at("notes")}</label>
                <textarea value={expenseForm.note} onChange={(e) => setExpenseForm({ ...expenseForm, note: e.target.value })} className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm" rows={2} />
              </div>
            </div>
            <button onClick={handleSaveExpense} className="w-full mt-4 bg-gradient-to-b from-teal-700 to-teal-900 text-white text-sm font-semibold rounded-xl py-2.5">
              {expenseEditingId ? at("saveChanges") : at("addExpense")}
            </button>
          </div>
        </div>
      )}

      {/* ---------- Accounts: treasury account form modal ---------- */}
      {showTreasuryAccountForm && (
        <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4" onClick={() => setShowTreasuryAccountForm(false)}>
          <div className="bg-white rounded-2xl w-full max-w-md p-5" onClick={(ev) => ev.stopPropagation()}>
            <div className="flex items-center justify-between mb-4">
              <h3 className="text-sm font-bold text-stone-800">{treasuryAccountEditingId ? at("editAccount") : at("addAccountTreasury")}</h3>
              <button onClick={() => setShowTreasuryAccountForm(false)} className="text-stone-400 hover:text-stone-700"><X size={18} /></button>
            </div>
            <div className="space-y-3">
              <div>
                <label className="text-xs text-stone-500 block mb-1">{at("accountName")}</label>
                <input value={treasuryForm.name} onChange={(e) => setTreasuryForm({ ...treasuryForm, name: e.target.value })} placeholder={at("accountNamePlaceholder")} className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm" />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">{at("type")}</label>
                <select value={treasuryForm.type} onChange={(e) => setTreasuryForm({ ...treasuryForm, type: e.target.value })} className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm">
                  {TREASURY_ACCOUNT_TYPES.map((t) => <option key={t.value} value={t.value}>{treasuryAccountTypeLabel(t.value)}</option>)}
                </select>
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">{at("openingBalance")}</label>
                <input type="number" value={treasuryForm.openingBalance} onChange={(e) => setTreasuryForm({ ...treasuryForm, openingBalance: e.target.value })} className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm" />
              </div>
            </div>
            <button onClick={handleSaveTreasuryAccount} className="w-full mt-4 bg-gradient-to-b from-teal-700 to-teal-900 text-white text-sm font-semibold rounded-xl py-2.5">
              {treasuryAccountEditingId ? at("saveChanges") : at("addAccount")}
            </button>
          </div>
        </div>
      )}

      {/* ---------- Accounts: manual treasury entry modal ---------- */}
      {showTreasuryEntryForm && (
        <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4" onClick={() => setShowTreasuryEntryForm(false)}>
          <div className="bg-white rounded-2xl w-full max-w-md p-5" onClick={(ev) => ev.stopPropagation()}>
            <div className="flex items-center justify-between mb-4">
              <h3 className="text-sm font-bold text-stone-800">{at("manualEntryTitle")}</h3>
              <button onClick={() => setShowTreasuryEntryForm(false)} className="text-stone-400 hover:text-stone-700"><X size={18} /></button>
            </div>
            <div className="space-y-3">
              <div className="flex gap-2">
                <button
                  onClick={() => setTreasuryEntryForm({ ...treasuryEntryForm, direction: "in", category: TREASURY_ENTRY_CATEGORIES_IN[0] })}
                  className={`flex-1 rounded-xl py-2 text-xs font-semibold border ${treasuryEntryForm.direction === "in" ? "bg-emerald-600 text-white border-emerald-600" : "bg-white border-stone-300 text-stone-500"}`}
                >
                  {at("directionIn")}
                </button>
                <button
                  onClick={() => setTreasuryEntryForm({ ...treasuryEntryForm, direction: "out", category: TREASURY_ENTRY_CATEGORIES_OUT[0] })}
                  className={`flex-1 rounded-xl py-2 text-xs font-semibold border ${treasuryEntryForm.direction === "out" ? "bg-red-600 text-white border-red-600" : "bg-white border-stone-300 text-stone-500"}`}
                >
                  {at("directionOut")}
                </button>
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">{at("colDate")}</label>
                <input type="date" value={treasuryEntryForm.date} onChange={(e) => setTreasuryEntryForm({ ...treasuryEntryForm, date: e.target.value })} className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm" />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">{at("colAccount")}</label>
                <select value={treasuryEntryForm.accountId} onChange={(e) => setTreasuryEntryForm({ ...treasuryEntryForm, accountId: e.target.value })} className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm">
                  <option value="">{at("selectAccount")}</option>
                  {treasuryAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
                </select>
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">{at("item")}</label>
                <select value={treasuryEntryForm.category} onChange={(e) => setTreasuryEntryForm({ ...treasuryEntryForm, category: e.target.value })} className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm">
                  {(treasuryEntryForm.direction === "in" ? TREASURY_ENTRY_CATEGORIES_IN : TREASURY_ENTRY_CATEGORIES_OUT).map((c) => (
                    <option key={c} value={c}>{treasuryEntryCategoryLabel(c)}</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">{at("amountEgp")}</label>
                <input type="number" value={treasuryEntryForm.amount} onChange={(e) => setTreasuryEntryForm({ ...treasuryEntryForm, amount: e.target.value })} className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm" />
              </div>
              <div>
                <label className="text-xs text-stone-500 block mb-1">{at("notes")}</label>
                <input value={treasuryEntryForm.note} onChange={(e) => setTreasuryEntryForm({ ...treasuryEntryForm, note: e.target.value })} className="w-full border border-stone-300 rounded-xl px-3 py-2 text-sm" />
              </div>
            </div>
            <button onClick={handleSaveTreasuryEntry} className="w-full mt-4 bg-gradient-to-b from-teal-700 to-teal-900 text-white text-sm font-semibold rounded-xl py-2.5">
              {at("saveEntry")}
            </button>
          </div>
        </div>
      )}

      {/* ---------- Accounts: supplier detail drawer ---------- */}
      {viewingSupplier && (
        <div className="fixed inset-0 bg-white z-40 overflow-y-auto">
          <div className="max-w-2xl mx-auto p-4 md:p-6">
            <div className="flex items-center justify-between mb-6">
              <h1 className="text-lg md:text-xl font-bold text-stone-900" style={{ fontFamily: "'Fraunces', serif" }}>{viewingSupplier}</h1>
              <button onClick={() => setViewingSupplier(null)} className="text-stone-400 hover:text-stone-700 p-1.5"><X size={18} /></button>
            </div>
            {(() => {
              const s = supplierLedger.find((x) => x.supplier === viewingSupplier) || { totalOwed: 0, paid: 0, balance: 0 };
              const bookings = acctBookings.filter((b) => b.supplier === viewingSupplier).sort((a, b) => (b.date || "").localeCompare(a.date || ""));
              const payments = supplierPayments.filter((p) => p.supplier === viewingSupplier).sort((a, b) => (b.date || "").localeCompare(a.date || ""));
              return (
                <>
                  <div className="grid grid-cols-3 gap-3 mb-6">
                    <div className="bg-stone-50 rounded-xl p-3 text-center">
                      <p className="text-[11px] text-stone-500 mb-1">{at("colTotalOwed")}</p>
                      <p className="font-bold text-stone-800">{fmt(s.totalOwed)}</p>
                    </div>
                    <div className="bg-stone-50 rounded-xl p-3 text-center">
                      <p className="text-[11px] text-stone-500 mb-1">{at("colPaid")}</p>
                      <p className="font-bold text-emerald-700">{fmt(s.paid)}</p>
                    </div>
                    <div className="bg-stone-50 rounded-xl p-3 text-center">
                      <p className="text-[11px] text-stone-500 mb-1">{at("colRemaining")}</p>
                      <p className={`font-bold ${s.balance > 0 ? "text-red-600" : "text-stone-500"}`}>{fmt(s.balance)}</p>
                    </div>
                  </div>

                  <div className="bg-teal-50/60 border border-teal-100 rounded-2xl p-4 mb-6">
                    <h3 className="text-xs font-bold text-teal-900 mb-3">{at("recordNewPayment")}</h3>
                    <div className="grid grid-cols-2 gap-2 mb-2">
                      <input type="date" value={supplierPaymentForm.date} onChange={(e) => setSupplierPaymentForm({ ...supplierPaymentForm, supplier: viewingSupplier, date: e.target.value })} className="border border-stone-300 rounded-lg px-2.5 py-2 text-sm" />
                      <input type="number" placeholder={at("colAmount")} value={supplierPaymentForm.amount} onChange={(e) => setSupplierPaymentForm({ ...supplierPaymentForm, supplier: viewingSupplier, amount: e.target.value })} className="border border-stone-300 rounded-lg px-2.5 py-2 text-sm" />
                      <select value={supplierPaymentForm.accountId} onChange={(e) => setSupplierPaymentForm({ ...supplierPaymentForm, supplier: viewingSupplier, accountId: e.target.value })} className="border border-stone-300 rounded-lg px-2.5 py-2 text-sm col-span-2">
                        <option value="">{at("payFrom")}</option>
                        {treasuryAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
                      </select>
                      <input placeholder={at("notesOptional")} value={supplierPaymentForm.note} onChange={(e) => setSupplierPaymentForm({ ...supplierPaymentForm, supplier: viewingSupplier, note: e.target.value })} className="border border-stone-300 rounded-lg px-2.5 py-2 text-sm col-span-2" />
                    </div>
                    <button onClick={handleSaveSupplierPayment} className="w-full bg-teal-800 hover:bg-teal-900 text-white text-xs font-semibold rounded-lg py-2">{at("recordPayment")}</button>
                  </div>

                  <h3 className="text-xs font-bold text-stone-600 mb-2">{at("paymentHistory")}</h3>
                  <div className="space-y-2 mb-6">
                    {payments.length === 0 ? (
                      <p className="text-xs text-stone-400">{at("noPaymentsRecorded")}</p>
                    ) : (
                      payments.map((p) => (
                        <div key={p.id} className="flex items-center justify-between bg-white border border-stone-200 rounded-xl px-3 py-2 text-sm">
                          <div>
                            <p className="font-semibold text-stone-800">{fmt(parseFloat(p.amount) || 0)} {acctCurrency}</p>
                            <p className="text-[11px] text-stone-400">{p.date ? formatDisplayDate(p.date) : "-"} · {treasuryAccounts.find((a) => a.id === p.accountId)?.name || "-"}{p.note ? ` · ${p.note}` : ""}</p>
                          </div>
                          <button onClick={() => handleDeleteSupplierPayment(p.id)} className="text-stone-400 hover:text-red-600"><Trash2 size={14} /></button>
                        </div>
                      ))
                    )}
                  </div>

                  <h3 className="text-xs font-bold text-stone-600 mb-2">{at("relatedBookings")}</h3>
                  <div className="space-y-2">
                    {bookings.length === 0 ? (
                      <p className="text-xs text-stone-400">{at("noBookings")}</p>
                    ) : (
                      bookings.map((b) => (
                        <div key={b.key} className="flex items-center justify-between bg-white border border-stone-200 rounded-xl px-3 py-2 text-sm">
                          <div>
                            <p className="font-semibold text-stone-800">{sectionLabel(b.section)} · {b.customers.join(", ") || "-"}</p>
                            <p className="text-[11px] text-stone-400">{b.date ? formatDisplayDate(b.date) : "-"}</p>
                          </div>
                          <p className="font-semibold text-stone-700">{fmt(b.net)} {acctCurrency}</p>
                        </div>
                      ))
                    )}
                  </div>
                </>
              );
            })()}
          </div>
        </div>
      )}

      {/* ---------- Accounts: customer detail drawer ---------- */}
      {viewingCustomer && (
        <div className="fixed inset-0 bg-white z-40 overflow-y-auto">
          <div className="max-w-2xl mx-auto p-4 md:p-6">
            <div className="flex items-center justify-between mb-6">
              <h1 className="text-lg md:text-xl font-bold text-stone-900" style={{ fontFamily: "'Fraunces', serif" }}>{viewingCustomer}</h1>
              <button onClick={() => setViewingCustomer(null)} className="text-stone-400 hover:text-stone-700 p-1.5"><X size={18} /></button>
            </div>
            {(() => {
              const c = customerLedger.find((x) => x.customer === viewingCustomer) || { totalDue: 0, paid: 0, balance: 0 };
              const bookings = acctBookings.filter((b) => b.customers.includes(viewingCustomer)).sort((a, b) => (b.date || "").localeCompare(a.date || ""));
              const payments = customerPayments.filter((p) => p.customer === viewingCustomer).sort((a, b) => (b.date || "").localeCompare(a.date || ""));
              return (
                <>
                  <div className="grid grid-cols-3 gap-3 mb-6">
                    <div className="bg-stone-50 rounded-xl p-3 text-center">
                      <p className="text-[11px] text-stone-500 mb-1">{at("colTotalDue")}</p>
                      <p className="font-bold text-stone-800">{fmt(c.totalDue)}</p>
                    </div>
                    <div className="bg-stone-50 rounded-xl p-3 text-center">
                      <p className="text-[11px] text-stone-500 mb-1">{at("colCollected")}</p>
                      <p className="font-bold text-emerald-700">{fmt(c.paid)}</p>
                    </div>
                    <div className="bg-stone-50 rounded-xl p-3 text-center">
                      <p className="text-[11px] text-stone-500 mb-1">{at("colRemaining")}</p>
                      <p className={`font-bold ${c.balance > 0 ? "text-red-600" : "text-stone-500"}`}>{fmt(c.balance)}</p>
                    </div>
                  </div>

                  <div className="bg-teal-50/60 border border-teal-100 rounded-2xl p-4 mb-6">
                    <h3 className="text-xs font-bold text-teal-900 mb-3">{at("recordNewCollection")}</h3>
                    <div className="grid grid-cols-2 gap-2 mb-2">
                      <input type="date" value={customerPaymentForm.date} onChange={(e) => setCustomerPaymentForm({ ...customerPaymentForm, customer: viewingCustomer, date: e.target.value })} className="border border-stone-300 rounded-lg px-2.5 py-2 text-sm" />
                      <input type="number" placeholder={at("colAmount")} value={customerPaymentForm.amount} onChange={(e) => setCustomerPaymentForm({ ...customerPaymentForm, customer: viewingCustomer, amount: e.target.value })} className="border border-stone-300 rounded-lg px-2.5 py-2 text-sm" />
                      <select value={customerPaymentForm.accountId} onChange={(e) => setCustomerPaymentForm({ ...customerPaymentForm, customer: viewingCustomer, accountId: e.target.value })} className="border border-stone-300 rounded-lg px-2.5 py-2 text-sm col-span-2">
                        <option value="">{at("collectInto")}</option>
                        {treasuryAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
                      </select>
                      <input placeholder={at("notesOptional")} value={customerPaymentForm.note} onChange={(e) => setCustomerPaymentForm({ ...customerPaymentForm, customer: viewingCustomer, note: e.target.value })} className="border border-stone-300 rounded-lg px-2.5 py-2 text-sm col-span-2" />
                    </div>
                    <button onClick={handleSaveCustomerPayment} className="w-full bg-teal-800 hover:bg-teal-900 text-white text-xs font-semibold rounded-lg py-2">{at("recordCollection")}</button>
                  </div>

                  <h3 className="text-xs font-bold text-stone-600 mb-2">{at("collectionHistory")}</h3>
                  <div className="space-y-2 mb-6">
                    {payments.length === 0 ? (
                      <p className="text-xs text-stone-400">{at("noCollectionsRecorded")}</p>
                    ) : (
                      payments.map((p) => (
                        <div key={p.id} className="flex items-center justify-between bg-white border border-stone-200 rounded-xl px-3 py-2 text-sm">
                          <div>
                            <p className="font-semibold text-stone-800">{fmt(parseFloat(p.amount) || 0)} {acctCurrency}</p>
                            <p className="text-[11px] text-stone-400">{p.date ? formatDisplayDate(p.date) : "-"} · {treasuryAccounts.find((a) => a.id === p.accountId)?.name || "-"}{p.note ? ` · ${p.note}` : ""}</p>
                          </div>
                          <button onClick={() => handleDeleteCustomerPayment(p.id)} className="text-stone-400 hover:text-red-600"><Trash2 size={14} /></button>
                        </div>
                      ))
                    )}
                  </div>

                  <h3 className="text-xs font-bold text-stone-600 mb-2">{at("relatedBookings")}</h3>
                  <div className="space-y-2">
                    {bookings.length === 0 ? (
                      <p className="text-xs text-stone-400">{at("noBookings")}</p>
                    ) : (
                      bookings.map((b) => (
                        <div key={b.key} className="flex items-center justify-between bg-white border border-stone-200 rounded-xl px-3 py-2 text-sm">
                          <div>
                            <p className="font-semibold text-stone-800">{sectionLabel(b.section)}</p>
                            <p className="text-[11px] text-stone-400">{b.date ? formatDisplayDate(b.date) : "-"}</p>
                          </div>
                          <p className="font-semibold text-stone-700">{fmt(b.sold / (b.customers.length || 1))} {acctCurrency}</p>
                        </div>
                      ))
                    )}
                  </div>
                </>
              );
            })()}
          </div>
        </div>
      )}
    </div>
  );
}
